import assert from 'node:assert/strict';
import { associateRows, normalizeMedicineProduct } from './src/app.js';

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
  hospital({ __rowId: 7, __os: '13157787', __osNormalizada: '1315778', __codBarra: '', __qtde: 250, __cliente: 'HOSPITAL SANTA HELENA', __paciente: 'VLADEMIR PERNIQUELLI', __medicamento: 'FAULDFLUOR 500MG', __principioAtivo: 'FAULDFLUOR 500MG' }),
];
const mandatoryControlRows = [
  control({ __rowId: 6, __os: '13157787', __osNormalizada: '1315778', __codBarra: '', __qtde: 200, __lote: '25D0738', __data: '01/06/2026', __unidadeDestino: 'HOSPITAL SANTA HELENA', __medicamento: 'FAULDFLUOR - 500 mg' }),
];
const mandatoryAssociations = associateRows(mandatoryHospitalRows, mandatoryControlRows, mandatoryValidations, 'Hospital Santa Helena');
const mandatoryAssociation = mandatoryAssociations.find((item) => item.hospitalRowId === 7 && item.controlRowId === 6);
assert.ok(mandatoryAssociation, 'associa a linha OTIMIZAÇÃO do controle sem CodBarra usando OS + medicamento normalizados');
assert.equal(mandatoryAssociation.qtdeOtimizada, 200, 'Qtde Otimizada vem da quantidade registrada na linha OTIMIZAÇÃO, limitada pela prescrição');
assert.equal(mandatoryAssociation.loteOtimizacao, '25D0738', 'Lote Otimização vem exatamente da linha OTIMIZAÇÃO do controle');
assert.equal(mandatoryAssociation.statusAssociacao, 'Parcialmente otimizado');
assert.equal(mandatoryAssociation.tipoMatch, 'MATCH OS + MEDICAMENTO');
assert.match(mandatoryAssociation.observacao, /linha OTIMIZAÇÃO do controle/);
assert.equal(mandatoryControlRows[0].__usedQuantity, 200);
assert.equal(mandatoryControlRows[0].__status, 'Consumido');
assert.ok(!mandatoryValidations.some((item) => item.Linha === 6 && item.Aba === 'CONTROLE DE MEDICAMENTOS' && item.Severidade === 'BLOQUEIO'), 'VALIDACAO não bloqueia a otimização sem CodBarra no controle');

console.log('association optimization tests passed');
