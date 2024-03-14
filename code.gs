function onInstall(e) {
  onOpen(e)
}

function onOpen(){
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Formatar CPFs selecionados', 'formatarIntervalo')
    .addToUi()
}

function clearCpfString(input){
  let strCPF = input.toString().replace(/[\D]/g, '')
  if(strCPF.length < 11) return '0'.repeat(11 - strCPF.length) + strCPF
  else return strCPF
}

function addCpfFormat(str){
  let start = str.substring(0,3)
  let middle = str.substring(3,6)
  let end = str.substring(6,9)
  let digits = str.substring(9,11)
  return `${start}.${middle}.${end}-${digits}`
}

function isValidCpf(str) {

  if (str.length > 11) return false

  const firstDigit = str[0]
  if (str.split('').every(digit => digit === firstDigit)) return false

  var Soma;
  var Resto;
  Soma = 0;

  for (i=1; i<=9; i++) Soma = Soma + parseInt(str.substring(i-1, i)) * (11 - i);
  Resto = (Soma * 10) % 11;

  if ((Resto == 10) || (Resto == 11))  Resto = 0;
  if (Resto != parseInt(str.substring(9, 10)) ) return false

  Soma = 0;
  for (i = 1; i <= 10; i++) Soma = Soma + parseInt(str.substring(i-1, i)) * (12 - i);
  Resto = (Soma * 10) % 11;

  if ((Resto == 10) || (Resto == 11))  Resto = 0;
  if (Resto != parseInt(str.substring(10, 11) ) ) return false
  
  return true
}

/**
 * Checa se o valor é um CPF válido e adiciona a formatação de CPF.
 * @customfunction
 */
function CPF(valor){
  if(valor === '') return valor

  let cpf = clearCpfString(valor)
  if(!isValidCpf(cpf)) return valor + ' (CPF inválido)'
  
  return addCpfFormat(cpf)
}

function formatarIntervalo(){
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getActiveSheet()
  const rng = sheet.getActiveRange()

  const values = rng.getValues()
  const mappedValues = values.map(row => row.map(cell => CPF(cell)))
  rng.setValues(mappedValues)

}
