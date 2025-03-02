// arquivo: backup.gs
// versão: 1.0
// autor: Juliano Ceconi

// Trigger para backup diário
function criarTriggerBackup() {
  ScriptApp.newTrigger('backupHistorico')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();
}