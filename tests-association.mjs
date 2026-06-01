import assert from 'node:assert/strict';
import { associateRows, formatLotWithValidity, normalizeMedicineBase, normalizeMedicineProduct, normalizeOS, readControlRows, summarizeAssociations } from './src/app.js';

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
assert.equal(fallbackControlRows[0].__validade, '30/04/2027', 'readControlRows lê Validade pela posição 10');
assert.equal(formatLotWithValidity(fallbackControlRows[0].__lote, fallbackControlRows[0].__validade), '25D0738 - Val.: 30/04/2027', 'lote e validade são combinados no formato do relatório');

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
assert.equal(mandatoryHospitalRows[0].__qtde - mandatoryAssociation.qtdeOtimizada, 100, 'Qtde Baixa esperada é 300 - 200 = 100');
assert.ok(mandatoryValidations.some((item) => item.Tipo === 'Divergência de lote como alerta não bloqueante'), 'lote diferente gera alerta não bloqueante');

console.log('association optimization tests passed');
