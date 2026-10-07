'use strict';
const Finance = MarionFinance;
let monthExpenses = [];
let expensesReady = false;
let expensesRequest = 0;
let expenseEditingId = null;
let expenseMonth = '';
let expenseBusy = false;
let excelLoading;

function initFinance() {
  $('expenseCategory').replaceChildren(...Finance.categories.map(category => {
    const option = document.createElement('option'); option.value = category; option.textContent = category; return option;
  }));
  $('expenseForm').addEventListener('submit', saveExpense);
  $('expenseOverlay').addEventListener('click', e => { if (e.target === e.currentTarget) closeExpenseForm(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeExpenseForm(); });
}
function financeStatus(message, error = false) {
  $('financeStatus').textContent = message;
  $('financeStatus').className = 'status' + (error ? ' error' : '');
}
function expenseError(error) {
  if (['42P01','PGRST205','42501'].includes(error.code)) return 'Le suivi des dépenses n’est pas encore activé. Exécute le script SQL fourni dans Supabase, puis recharge cette page.';
  return error.message || 'Connexion interrompue. Réessaie.';
}
function expensesQuery(client, start, end) {
  return () => client.from('marion_expenses').select('id,expense_date,category,label,amount').gte('expense_date', start).lt('expense_date', end).order('expense_date').order('id');
}
async function loadMonthExpenses() {
  const request = ++expensesRequest;
  expensesReady = false;
  monthExpenses = [];
  renderExpenses();
  updateFinanceSummary();
  $('financeMonth').textContent = new Intl.DateTimeFormat('fr-FR', {month:'long',year:'numeric'}).format(viewDate);
  $('exportExcelBtn').textContent = 'Exporter Excel ' + viewDate.getFullYear();
  financeStatus('Chargement des dépenses…');
  const {start,end} = monthBounds();
  try {
    const rows = await Finance.fetchAll(expensesQuery(db, start, end));
    if (request !== expensesRequest) return;
    monthExpenses = rows;
    expensesReady = true;
    financeStatus(rows.length + ' dépense(s) enregistrée(s) pour ce mois.');
  } catch (error) {
    if (request !== expensesRequest) return;
    financeStatus(expenseError(error), true);
  }
  renderExpenses();
  updateFinanceSummary();
}
function updateFinanceSummary() {
  $('sumExpenses').textContent = expensesReady ? money(Finance.sum(monthExpenses, 'amount')) : 'Indisponible';
  $('sumResult').textContent = expensesReady && monthRevenueReady ? money(Finance.summary(workDays, monthExpenses).result) : 'Indisponible';
}
function renderExpenses() {
  const box = $('expenseList');
  box.replaceChildren();
  $('addExpenseBtn').disabled = !expensesReady;
  if (!expensesReady) { box.textContent = 'Les dépenses ne sont pas encore disponibles.'; return; }
  if (!monthExpenses.length) { box.textContent = 'Aucune dépense pour ce mois.'; return; }
  monthExpenses.forEach(entry => {
    const item = document.createElement('div'); item.className = 'expenseItem';
    const title = document.createElement('strong'); title.textContent = entry.label;
    const meta = document.createElement('div'); meta.className = 'small';
    meta.textContent = parseIsoDate(entry.expense_date).toLocaleDateString('fr-FR') + ' · ' + entry.category + ' · ' + money(entry.amount);
    const button = document.createElement('button'); button.className = 'light'; button.textContent = 'Modifier';
    button.setAttribute('aria-label', 'Modifier ' + entry.label);
    button.onclick = () => openExpenseForm(entry);
    item.append(title, meta, button); box.append(item);
  });
}
function openExpenseForm(entry = null, fromSelectedDay = false) {
  if (!expensesReady) { alert('Attends le chargement des dépenses du mois.'); return; }
  expenseEditingId = entry?.id || null;
  expenseMonth = monthBounds().start.slice(0, 7);
  const first = expenseMonth + '-01';
  const last = isoLocal(new Date(viewDate.getFullYear(), viewDate.getMonth() + 1, 0));
  const defaultDate = fromSelectedDay && selectedDate.startsWith(expenseMonth) ? selectedDate : (todayIso().startsWith(expenseMonth) ? todayIso() : first);
  $('expenseDate').min = first; $('expenseDate').max = last;
  $('expenseDate').value = entry?.expense_date || defaultDate;
  $('expenseCategory').value = entry?.category || Finance.categories[0];
  $('expenseLabel').value = entry?.label || '';
  $('expenseAmount').value = entry?.amount ?? '';
  $('expenseFormTitle').textContent = (entry ? 'Modifier une dépense — ' : 'Ajouter une dépense — ') + $('financeMonth').textContent;
  $('deleteExpenseBtn').hidden = !entry;
  $('expenseFormStatus').textContent = '';
  closeWorkDayForm();
  $('expenseOverlay').classList.add('open');
  $('expenseLabel').focus();
}
function closeExpenseForm() {
  if (expenseBusy) return;
  $('expenseOverlay').classList.remove('open');
}
function expenseSaving(busy) {
  expenseBusy = busy;
  $('expenseFields').disabled = busy;
  $('closeExpenseBtn').disabled = busy;
}
async function saveExpense(event) {
  event.preventDefault();
  if (expenseBusy) return;
  expenseSaving(true);
  try {
    const payload = Finance.validateExpense({ expense_date: $('expenseDate').value, category: $('expenseCategory').value, label: $('expenseLabel').value, amount: $('expenseAmount').value }, expenseMonth);
    const query = expenseEditingId ? db.from('marion_expenses').update(payload).eq('id', expenseEditingId) : db.from('marion_expenses').insert(payload);
    const { error } = await query.select('id').single();
    if (error) throw error;
    expenseSaving(false);
    closeExpenseForm();
    await loadMonthExpenses();
  } catch(error) { $('expenseFormStatus').textContent = expenseError(error); }
  finally { expenseSaving(false); }
}
async function deleteExpense() {
  if (expenseBusy || !expenseEditingId || !confirm('Supprimer cette dépense ?')) return;
  expenseSaving(true);
  try {
    const {error} = await db.from('marion_expenses').delete().eq('id', expenseEditingId).select('id').single();
    if (error) throw error;
    expenseSaving(false); closeExpenseForm();
    await loadMonthExpenses();
  } catch(error) { $('expenseFormStatus').textContent = expenseError(error); }
  finally { expenseSaving(false); }
}
function loadExcelLibrary() {
  if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
  if (excelLoading) return excelLoading;
  excelLoading = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
    const timer = setTimeout(() => { script.remove(); excelLoading = null; reject(new Error('Le téléchargement de l’outil Excel a expiré. Réessaie.')); }, 20000);
    script.onload = () => { clearTimeout(timer); if (window.ExcelJS) resolve(window.ExcelJS); else { excelLoading = null; reject(new Error('Outil Excel indisponible.')); } };
    script.onerror = () => { clearTimeout(timer); script.remove(); excelLoading = null; reject(new Error('Impossible de charger l’outil Excel. Vérifie ta connexion Internet.')); };
    document.head.append(script);
  });
  return excelLoading;
}
async function exportAnnualExcel() {
  const button = $('exportExcelBtn');
  if (button.disabled) return;
  const year = viewDate.getFullYear();
  button.disabled = true;
  financeStatus('Préparation du fichier Excel ' + year + '…');
  try {
    const start = `${year}-01-01`, end = `${year+1}-01-01`;
    const [ExcelJS, work, expenses, allTours] = await Promise.all([
      loadExcelLibrary(),
      Finance.fetchAll(() => db.from('work_days').select('id,work_date,tour_id,revenue,is_paid').gte('work_date', start).lt('work_date', end).order('work_date').order('id')),
      Finance.fetchAll(expensesQuery(db, start, end)),
      Finance.fetchAll(() => db.from('tours').select('id,name').order('id'))
    ]);
    const workbook = Finance.buildWorkbook(ExcelJS, year, work, expenses, allTours);
    const contributionDefaults = MarionCotisations.prefill(work, expenses);
    let contributions;
    try {
      contributions = await MarionCotisations.estimate({year,...contributionDefaults});
    } catch(error) {
      contributions = {error:error.name === 'AbortError' ? 'Le simulateur URSSAF ne répond pas. Réexporter plus tard.' : error.message};
    }
    MarionCotisations.appendWorkbook(workbook,year,contributionDefaults,contributions,Finance.categories,Finance.months);
    const Charts = await loadChartModule();
    const chartPlan = Charts.appendWorkbook(workbook,year,work,allTours,Finance);
    const JSZip = await loadChartLibrary();
    const bytes = await Charts.writeBuffer(workbook,chartPlan,JSZip);
    const url = URL.createObjectURL(new Blob([bytes], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
    const a = document.createElement('a'); a.href = url; a.download = `Tournee_Marion_${year}.xlsx`;
    document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    financeStatus('Fichier Excel ' + year + ' téléchargé : 12 mois, total annuel et graphiques.' + (contributions.error ? ' Estimation sociale indisponible : voir le classeur.' : ''));
  } catch(error) { financeStatus('Export annulé : ' + expenseError(error), true); }
  finally { button.disabled = false; }
}
let chartModuleLoading;
function loadChartModule() {
  if (window.MarionCharts) return Promise.resolve(window.MarionCharts);
  if (chartModuleLoading) return chartModuleLoading;
  chartModuleLoading = new Promise((resolve,reject)=>{
    const script=document.createElement('script');
    script.src=new URL('finance-charts.js?v=20261007-2',document.baseURI).href;
    const fail=()=>{clearTimeout(timer);script.remove();chartModuleLoading=null;reject(new Error('Le module des graphiques ne peut pas être chargé. Vérifie la connexion puis réessaie.'));};
    const timer=setTimeout(fail,20000);
    script.onload=()=>{if(!window.MarionCharts)return fail();clearTimeout(timer);resolve(window.MarionCharts);};
    script.onerror=fail;document.head.append(script);
  });
  return chartModuleLoading;
}
let chartLibraryLoading;
function loadChartLibrary() {
  if (window.JSZip) return Promise.resolve(window.JSZip);
  if (chartLibraryLoading) return chartLibraryLoading;
  chartLibraryLoading = new Promise((resolve,reject)=>{
    const script=document.createElement('script');
    script.src='https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js';
    const fail=()=>{clearTimeout(timer);script.remove();chartLibraryLoading=null;reject(new Error('Impossible de charger les graphiques Excel. Réessaie.'));};
    const timer=setTimeout(fail,20000);
    script.onload=()=>{if(!window.JSZip)return fail();clearTimeout(timer);resolve(window.JSZip);};
    script.onerror=fail;document.head.append(script);
  });
  return chartLibraryLoading;
}
