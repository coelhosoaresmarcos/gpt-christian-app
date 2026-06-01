import assert from 'node:assert/strict';
import { associateRows, normalizeMedicineProduct, normalizeOS, readControlRows, summarizeAssociations } from './src/app.js';


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
  const row = {
    __rowId: overrides.__rowId,
    __data: overrides.__data ?? '01/01/2026',
    __cliente: overrides.__cliente ?? 'Hospital Teste',
    __paciente: overrides.__paciente ?? 'Paciente',
    __medicamento: overrides.__medicamento ?? 'FAULDFLUOR 500MG',
    __principioAtivo: overrides.__principioAtivo ?? 'FAULDFLUOR 500MG',
    __codBarra: overrides.__codBarra ?? '',
    __qtde: overrides.__qtde,
    __lote: 'LOTE-H',
    __os: overrides.__os,
    __osNormalizada: overrides.__osNormalizada,
    __key: `${overrides.__osNormalizada}|${overrides.__codBarra ?? ''}`,
    __isExcluded: false,
  };
  if (!overrides.__osNormalizada || !overrides.__codBarra) row.__key = '';
  return row;
}

function control(overrides) {
  const row = {
    __rowId: overrides.__rowId,
    __data: overrides.__data ?? '31/12/2025',
    __os: overrides.__os,
    __osNormalizada: overrides.__osNormalizada,
    __medicamento: overrides.__medicamento ?? 'FAULDFLUOR - 500 mg',
    __medicamentoNormalizado: normalizeMedicineProduct(overrides.__medicamento ?? 'FAULDFLUOR - 500 mg'),
    __qtde: overrides.__qtde,
    __lote: overrides.__lote,
    __motivo: overrides.__motivo ?? 'OTIMIZAÇÃO',
    __unidadeDestino: overrides.__unidadeDestino ?? 'Hospital Teste',
    __codBarra: overrides.__codBarra ?? '',
    __key: `${overrides.__osNormalizada}|${overrides.__codBarra ?? ''}`,
    __status: 'Não avaliado como otimização',
    __observacao: '',
    __tipoMatch: '',
    __confianca: '',
    __hospitalAssociado: '',
    __dataAssociacao: '',
    __used: false,
  };
  if (!overrides.__osNormalizada || !overrides.__codBarra) row.__key = '';
  return row;
}

assert.equal(normalizeMedicineProduct('FAULDFLUOR - 500 mg'), 'FAULDFLUOR 500MG');

const fallbackControlRows = readControlRows(fakeSheet([
  ['Tipo/Motivo', 'OS', 'Data', 'Unidade Origem', 'Unidade Destino', 'Paciente', 'Medicamento', 'Quantidade', 'Lote', 'Validade', 'Laboratório'],
  ['OTIMIZAÇÃO', '13157787', '01/06/2026', 'CENTRO MEDICO PITANGUEIRAS', 'HOSPITAL SANTA HELENA', 'VLADEMIR PERNIQUELLI', 'FAULDFLUOR - 500 mg', 200, '25D0738', '30/04/2027', 'LIBBS'],
]), []);
assert.equal(fallbackControlRows[0].__os, '13157787', 'readControlRows lê OS pela posição 2 quando necessário');
assert.equal(fallbackControlRows[0].__osNormalizada, '13157787', 'OS lida no controle permanece com todos os dígitos');
assert.equal(fallbackControlRows[0].__medicamento, 'FAULDFLUOR - 500 mg', 'readControlRows lê medicamento pela posição 7');
assert.equal(fallbackControlRows[0].__qtde, 200, 'readControlRows lê quantidade pela posição 8');
assert.equal(fallbackControlRows[0].__lote, '25D0738', 'readControlRows lê lote pela posição 9');

const positionalFallbackControlRows = readControlRows(fakeSheet([
  ['Motivo sem nome oficial', 'Identificador atendimento', 'Quando', 'Origem remessa', 'Destino remessa', 'Nome beneficiario', 'Item solicitado', 'Saldo remessa', 'Numero lote', 'Vence em', 'Fabricante'],
  ['OTIMIZAÇÃO', '13157787', '01/06/2026', 'CENTRO MEDICO PITANGUEIRAS', 'HOSPITAL SANTA HELENA', 'VLADEMIR PERNIQUELLI', 'FAULDFLUOR - 500 mg', 200, '25D0738', '30/04/2027', 'LIBBS'],
]), []);
assert.equal(positionalFallbackControlRows[0].__motivo, 'OTIMIZAÇÃO', 'fallback por posição lê Tipo/Motivo na coluna 1');
assert.equal(positionalFallbackControlRows[0].__os, '13157787', 'fallback por posição lê OS na coluna 2');
assert.equal(positionalFallbackControlRows[0].__data, '01/06/2026', 'fallback por posição lê Data na coluna 3');
assert.equal(positionalFallbackControlRows[0].__unidadeOrigem, 'CENTRO MEDICO PITANGUEIRAS', 'fallback por posição lê Unidade Origem na coluna 4');
assert.equal(positionalFallbackControlRows[0].__unidadeDestino, 'HOSPITAL SANTA HELENA', 'fallback por posição lê Unidade Destino na coluna 5');
assert.equal(positionalFallbackControlRows[0].__paciente, 'VLADEMIR PERNIQUELLI', 'fallback por posição lê Paciente na coluna 6');
assert.equal(positionalFallbackControlRows[0].__medicamento, 'FAULDFLUOR - 500 mg', 'fallback por posição lê Medicamento na coluna 7');
assert.equal(positionalFallbackControlRows[0].__qtde, 200, 'fallback por posição lê Quantidade na coluna 8');
assert.equal(positionalFallbackControlRows[0].__lote, '25D0738', 'fallback por posição lê Lote na coluna 9');
assert.equal(positionalFallbackControlRows[0].__validade, '30/04/2027', 'fallback por posição lê Validade na coluna 10');
assert.equal(positionalFallbackControlRows[0].__laboratorio, 'LIBBS', 'fallback por posição lê Laboratório na coluna 11');

assert.equal(normalizeOS('13157787.0'), '13157787', 'normalizeOS mantém todos os 8 dígitos da OS real e remove apenas o decimal do Excel');

const validations = [];
const hospitalRows = [
  hospital({ __rowId: 2, __os: '1234567', __osNormalizada: '1234567', __codBarra: '7891000', __qtde: 5 }),
  hospital({ __rowId: 3, __os: '1234567', __osNormalizada: '1234567', __codBarra: '7892000', __qtde: 3 }),
  hospital({ __rowId: 4, __os: '7654321', __osNormalizada: '7654321', __codBarra: '7891000', __qtde: 2 }),
  hospital({ __rowId: 5, __os: '9999999', __osNormalizada: '9999999', __codBarra: '7893000', __qtde: 4 }),
  hospital({ __rowId: 6, __os: '9999999', __osNormalizada: '9999999', __codBarra: '7893000', __qtde: 4 }),
];
const controlRows = [
  control({ __rowId: 2, __os: '1234567', __osNormalizada: '1234567', __codBarra: '7891000', __qtde: 5, __lote: 'OK' }),
  control({ __rowId: 3, __os: '1234567', __osNormalizada: '1234567', __codBarra: '0000000', __qtde: 10, __lote: 'OS-ONLY-BLOCKED' }),
  control({ __rowId: 4, __os: '0000000', __osNormalizada: '0000000', __codBarra: '7891000', __qtde: 10, __lote: 'BARCODE-ONLY-BLOCKED' }),
  control({ __rowId: 5, __os: '9999999', __osNormalizada: '9999999', __codBarra: '7893000', __qtde: 6, __lote: 'SALDO' }),
];

const associations = associateRows(hospitalRows, controlRows, validations, 'Hospital Teste');

const row2Used = associations.filter((item) => item.hospitalRowId === 2).reduce((sum, item) => sum + item.qtdeUsada, 0);
assert.equal(row2Used, 5, 'associa quando OS + CodBarra são iguais e medicamento é compatível');
assert.equal(controlRows[0].__usedQuantity, 5, 'consome quantidade do controle correspondente');

const row3Used = associations.filter((item) => item.hospitalRowId === 3).reduce((sum, item) => sum + item.qtdeUsada, 0);
assert.equal(row3Used, 0, 'não associa quando OS é igual, mas CodBarra é diferente e o controle possui CodBarra');
assert.equal(controlRows[1].__usedQuantity, 0, 'não usa saldo com apenas OS igual quando há CodBarra divergente');
assert.ok(validations.some((item) => item.Tipo === 'Correspondência por OS, mas CodBarra diferente'));

const row4Used = associations.filter((item) => item.hospitalRowId === 4).reduce((sum, item) => sum + item.qtdeUsada, 0);
assert.equal(row4Used, 0, 'não associa quando CodBarra é igual, mas OS é diferente');
assert.equal(controlRows[2].__usedQuantity, 0, 'não usa saldo com apenas CodBarra igual');
assert.ok(validations.some((item) => item.Tipo === 'Correspondência por CodBarra, mas OS diferente'));

const row5Used = associations.filter((item) => item.hospitalRowId === 5).reduce((sum, item) => sum + item.qtdeUsada, 0);
const row6Used = associations.filter((item) => item.hospitalRowId === 6).reduce((sum, item) => sum + item.qtdeUsada, 0);
assert.equal(row5Used, 4, 'primeira linha consome saldo disponível');
assert.equal(row6Used, 2, 'segunda linha usa somente saldo remanescente');
assert.equal(controlRows[3].__usedQuantity, 6, 'saldo do controle não é usado duas vezes');
assert.equal(controlRows[3].__remaining, 0, 'saldo remanescente chega a zero');

const mandatoryValidations = [];
const mandatoryHospitalRows = [
  hospital({ __rowId: 7, __os: '13157787', __osNormalizada: normalizeOS('13157787'), __codBarra: '', __qtde: 300, __cliente: 'HOSPITAL SANTA HELENA', __paciente: 'VLADEMIR PERNIQUELLI', __medicamento: 'FAULDFLUOR - 500 mg', __principioAtivo: 'FAULDFLUOR - 500 mg' }),
];
const mandatoryControlRows = [
  control({ __rowId: 6, __os: '13157787', __osNormalizada: normalizeOS('13157787'), __codBarra: '', __qtde: 200, __lote: '25D0738', __data: '01/06/2026', __unidadeDestino: 'HOSPITAL SANTA HELENA', __medicamento: 'FAULDFLUOR - 500 mg' }),
];
const mandatoryAssociations = associateRows(mandatoryHospitalRows, mandatoryControlRows, mandatoryValidations, 'Hospital Santa Helena');
const mandatoryAssociation = mandatoryAssociations.find((item) => item.hospitalRowId === 7 && item.controlRowId === 6);
assert.ok(mandatoryAssociation, 'associa a linha OTIMIZAÇÃO do controle sem CodBarra usando OS + medicamento normalizados');
assert.equal(mandatoryHospitalRows[0].__osNormalizada, '13157787', 'OS 13157787 não pode ser truncada');
assert.equal(mandatoryControlRows[0].__osNormalizada, '13157787', 'OS 13157787 do controle não pode ser truncada');
assert.equal(mandatoryAssociation.qtdeOtimizada, 200, 'Qtde Otimizada vem da quantidade registrada na linha OTIMIZAÇÃO, limitada pela prescrição');
assert.equal(mandatoryAssociation.loteOtimizacao, '25D0738', 'Lote Otimização vem exatamente da linha OTIMIZAÇÃO do controle');
assert.equal(mandatoryAssociation.statusAssociacao, 'Parcialmente otimizado');
const mandatoryReportSummary = summarizeAssociations(mandatoryAssociations).get(7);
assert.deepEqual(mandatoryReportSummary.lotes, ['25D0738'], 'RELATORIO/BAIXAR usa o lote 25D0738 resumido para a otimização');
assert.equal(mandatoryReportSummary.status, 'Parcialmente otimizado', 'RELATORIO/BAIXAR mantém status Parcialmente otimizado');
assert.equal(mandatoryHospitalRows[0].__qtde - mandatoryAssociation.qtdeOtimizada, 100, 'Qtde Baixa esperada é 300 - 200 = 100');
assert.equal(mandatoryAssociation.tipoMatch, 'MATCH OS + MEDICAMENTO');
assert.match(mandatoryAssociation.observacao, /linha OTIMIZAÇÃO do controle/);
assert.equal(mandatoryControlRows[0].__usedQuantity, 200);
assert.equal(mandatoryControlRows[0].__status, 'Consumido');
assert.ok(!mandatoryValidations.some((item) => item.Linha === 6 && item.Aba === 'CONTROLE DE MEDICAMENTOS' && item.Severidade === 'BLOQUEIO'), 'VALIDACAO não bloqueia a otimização sem CodBarra no controle');

const destinationValidations = [];
const destinationHospitalRows = [
  hospital({ __rowId: 8, __os: '13157787', __osNormalizada: normalizeOS('13157787'), __codBarra: '', __qtde: 300, __cliente: 'HOSPITAL SANTA HELENA', __medicamento: 'FAULDFLUOR - 500 mg', __principioAtivo: 'FAULDFLUOR - 500 mg' }),
];
const destinationControlRows = [
  control({ __rowId: 8, __os: '13157787', __osNormalizada: normalizeOS('13157787'), __codBarra: '', __qtde: 200, __lote: '25D0738', __unidadeDestino: 'HOSPITAL SANTA HELENA', __medicamento: 'FAULDFLUOR - 500 mg' }),
];
const destinationAssociations = associateRows(destinationHospitalRows, destinationControlRows, destinationValidations, 'Santa Helena São Bernardo');
assert.ok(destinationAssociations.some((item) => item.hospitalRowId === 8 && item.controlRowId === 8), 'Unidade de Destino parcial/divergente não bloqueia associação por OS + medicamento');

console.log('association optimization tests passed');
