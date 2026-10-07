(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.MarionFinance = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const categories = ['Prélèvement salaire', 'Essence', 'Réparations', 'Leasing', 'Assurance', 'Cotisations', 'URSSAF', 'CARPIMKO', 'Matériel', 'Autre'];
  const months = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'];
  const euro = '#,##0.00 "€";[Red](#,##0.00) "€";0.00 "€"';
  function cents(value) {
    const n = Number(value);
    if (value == null || value === '' || !Number.isFinite(n) || n < 0 || n > 999999999.99) throw new Error('Montant invalide.');
    return Math.round((n + Number.EPSILON) * 100);
  }
  function sum(rows, field) { return rows.reduce((n, row) => n + cents(row[field]), 0) / 100; }
  function summary(work, expenses) {
    const revenue = sum(work, 'revenue');
    const deductions = sum(expenses, 'amount');
    return { revenue, deductions, result: Math.round((revenue - deductions) * 100) / 100 };
  }
  // Request a small page and advance by its actual size: also works with a lower server row cap.
  async function fetchAll(makeQuery) {
    const rows = [];
    for (;;) {
      const { data, error } = await makeQuery().range(rows.length, rows.length + 199);
      if (error) throw error;
      if (!Array.isArray(data)) throw new Error('Réponse de données incomplète.');
      if (!data.length) return rows;
      rows.push(...data);
    }
  }
  function validateExpense(raw, month) {
    const date = String(raw.expense_date || '');
    const parsed = new Date(date + 'T12:00:00Z');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !Number.isFinite(parsed.getTime()) || parsed.toISOString().slice(0, 10) !== date || date.slice(0, 7) !== month) {
      throw new Error('Choisis une date valide dans le mois indiqué.');
    }
    if (!categories.includes(raw.category)) throw new Error('Choisis une catégorie.');
    const amount = cents(raw.amount);
    if (amount <= 0) throw new Error('Le montant doit être supérieur à zéro.');
    const label = String(raw.label || '').trim();
    if (!label || label.length > 200) throw new Error('Saisis un libellé de 1 à 200 caractères.');
    return { expense_date: date, category: raw.category, label, amount: amount / 100 };
  }
  function buildWorkbook(ExcelJS, year, workDays, expenses, tours) {
    const book = new ExcelJS.Workbook();
    book.creator = 'Tournée Marion';
    book.calcProperties.fullCalcOnLoad = true;
    const names = new Map(tours.map(t => [t.id, t.name]));
    const totals = [];
    function setup(name, title) {
      const sheet = book.addWorksheet(name, {
        views: [{ showGridLines: false }],
        pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
      });
      sheet.columns = [{ width: 42 }, { width: 30 }, { width: 45 }, { width: 22 }, { width: 22 }];
      sheet.getCell('A2').value = title;
      sheet.getCell('A2').font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FF1769AA' } };
      sheet.getCell('A3').value = 'Toutes les tournées • Montants en euros';
      sheet.getCell('A3').font = { name: 'Arial', size: 10, italic: true, color: { argb: 'FF667085' } };
      return sheet;
    }
    function formula(sheet, address, formula, result) {
      const cell = sheet.getCell(address);
      cell.value = { formula, result };
      cell.numFmt = euro;
    }
    function percentage(sheet, address, numerator, denominator, value, revenue) {
      const cell = sheet.getCell(address);
      cell.value = { formula: `IF(${denominator}=0,"—",${numerator}/${denominator})`, result: revenue === 0 ? '—' : value / revenue };
      cell.numFmt = '0.0%';
      cell.alignment = { horizontal: 'right', vertical: 'middle' };
    }
    function heading(sheet, row, values) {
      sheet.getRow(row).values = values;
      sheet.getRow(row).height = 25;
      values.forEach((_, i) => {
        const cell = sheet.getCell(row, i + 1);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1769AA' } };
        cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      });
    }
    function finish(sheet) {
      sheet.eachRow(row => {
        if (!row.height) row.height = 22;
        row.eachCell(cell => {
          if (!cell.font) cell.font = { name: 'Arial', size: 11 };
          if (!cell.alignment) cell.alignment = { vertical: 'middle', wrapText: true };
        });
      });
    }
    for (let m = 0; m < 12; m++) {
      const key = `${year}-${String(m + 1).padStart(2, '0')}`;
      const work = workDays.filter(x => x.work_date.slice(0, 7) === key).sort((a,b) => a.work_date.localeCompare(b.work_date));
      const costs = expenses.filter(x => x.expense_date.slice(0, 7) === key).sort((a,b) => a.expense_date.localeCompare(b.expense_date));
      const total = summary(work, costs);
      total.remuneration = sum(costs.filter(x => x.category === 'Prélèvement salaire'), 'amount');
      total.otherExpenses = Math.round((total.deductions - total.remuneration) * 100) / 100;
      totals.push(total);
      const sheet = setup(months[m], `${months[m]} ${year}`);
      sheet.getCell('A5').value = 'Chiffre d’affaires';
      sheet.getCell('A6').value = 'Dépenses et prélèvements à déduire';
      sheet.getCell('A7').value = 'Résultat mensuel';
      sheet.getCell('A9').value = 'CA des journées saisies, payées ou non.';
      sheet.getCell('A10').value = 'Résultat = CA − dépenses, salaire inclus.';
      heading(sheet, 12, ['Catégorie', 'Montant à déduire', '% du CA du mois']);
      categories.forEach((category, i) => { sheet.getCell(i + 13, 1).value = category; });
      const expenseHeader = 13 + categories.length + 2;
      heading(sheet, expenseHeader, ['Date', 'Catégorie', 'Libellé', 'Dépense (€)']);
      const expenseStart = expenseHeader + 1;
      costs.forEach((x, i) => {
        const row = sheet.getRow(expenseStart + i);
        row.values = [new Date(x.expense_date + 'T00:00:00Z'), x.category, x.label, cents(x.amount) / 100];
        row.getCell(1).numFmt = 'dd/mm/yyyy';
        row.getCell(4).numFmt = euro;
        row.height = Math.max(24, Math.ceil(String(x.label).length / 42) * 15);
      });
      const expenseEnd = expenseStart + Math.max(costs.length, 1) - 1;
      if (!costs.length) sheet.getCell(expenseStart, 3).value = 'Aucune dépense enregistrée';
      categories.forEach((category, i) => {
        const amount = sum(costs.filter(x => x.category === category), 'amount');
        formula(sheet, `B${13 + i}`, `SUMIF(B${expenseStart}:B${expenseEnd},A${13 + i},D${expenseStart}:D${expenseEnd})`, amount);
        percentage(sheet, `C${13+i}`, `B${13+i}`, '$B$5', amount, total.revenue);
      });
      const workHeader = expenseEnd + 3;
      heading(sheet, workHeader, ['Date', 'Tournée', 'Paiement', 'Chiffre d’affaires (€)']);
      const workStart = workHeader + 1;
      work.forEach((x, i) => {
        const row = sheet.getRow(workStart + i);
        row.values = [new Date(x.work_date + 'T00:00:00Z'), names.get(x.tour_id) || 'Tournée archivée', x.is_paid ? 'Payée' : 'Non payée', cents(x.revenue) / 100];
        row.getCell(1).numFmt = 'dd/mm/yyyy';
        row.getCell(4).numFmt = euro;
      });
      const workEnd = workStart + Math.max(work.length, 1) - 1;
      if (!work.length) sheet.getCell(workStart, 2).value = 'Aucune journée enregistrée';
      formula(sheet, 'B5', `SUM(D${workStart}:D${workEnd})`, total.revenue);
      formula(sheet, 'B6', `SUM(D${expenseStart}:D${expenseEnd})`, total.deductions);
      formula(sheet, 'B7', 'B5-B6', total.result);
      sheet.getCell('C4').value = '% du CA du mois';
      percentage(sheet, 'C6', 'B6', '$B$5', total.deductions, total.revenue);
      sheet.mergeCells('A24:D24');
      sheet.getCell('A24').value = '— : pourcentage non calculable lorsque le CA est nul.';
      sheet.getCell('A24').font = { name: 'Arial', size: 10, italic: true, color: { argb: 'FF667085' } };
      ['A7','B7'].forEach(a => { sheet.getCell(a).font = { name: 'Arial', size: 12, bold: true }; sheet.getCell(a).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } }; });
      sheet.views = [{ showGridLines: false, state: 'frozen', ySplit: 7 }];
      finish(sheet);
    }
    const annual = setup('Total annuel', `Bilan annuel ${year}`);
    annual.getCell('A5').value = 'Chiffre d’affaires annuel';
    annual.columns = [{width:42},{width:24},{width:24},{width:18},{width:24},{width:18},{width:24}];
    annual.getCell('A6').value = 'Dépenses hors rémunération';
    annual.getCell('A7').value = 'Résultat annuel';
    annual.getCell('A8').value = 'Rémunération (prélèvements salaire)';
    annual.getCell('C4').value = '% du CA annuel';
    heading(annual, 10, ['Mois', 'Chiffre d’affaires (€)', 'Rémunération (€)', 'Rémunération / CA', 'Dépenses hors rémunération (€)', 'Dépenses / CA', 'Résultat (€)']);
    annual.getRow(10).height = 44;
    months.forEach((month, i) => {
      const r = i + 11;
      annual.getCell(r, 1).value = month;
      const t = totals[i];
      formula(annual, `B${r}`, `'${month}'!B5`, t.revenue);
      formula(annual, `C${r}`, `'${month}'!B${13+categories.indexOf('Prélèvement salaire')}`, t.remuneration);
      percentage(annual, `D${r}`, `C${r}`, `B${r}`, t.remuneration, t.revenue);
      formula(annual, `E${r}`, `'${month}'!B6-C${r}`, t.otherExpenses);
      percentage(annual, `F${r}`, `E${r}`, `B${r}`, t.otherExpenses, t.revenue);
      formula(annual, `G${r}`, `B${r}-C${r}-E${r}`, t.result);
    });
    annual.getCell('A23').value = 'Total des 12 mois';
    ['B','C','E','G'].forEach((col, i) => formula(annual, col+'23', `SUM(${col}11:${col}22)`, Math.round(totals.reduce((s,t) => s + [t.revenue,t.remuneration,t.otherExpenses,t.result][i] * 100,0)) / 100));
    const annualRevenue = annual.getCell('B23').result ?? 0;
    const remuneration = annual.getCell('C23').result ?? 0;
    const otherExpenses = annual.getCell('E23').result ?? 0;
    percentage(annual, 'D23', 'C23', 'B23', remuneration, annualRevenue);
    percentage(annual, 'F23', 'E23', 'B23', otherExpenses, annualRevenue);
    formula(annual, 'B5', 'B23', annualRevenue);
    formula(annual, 'B6', 'E23', otherExpenses);
    formula(annual, 'B8', 'C23', remuneration);
    formula(annual, 'B7', 'B5-B6-B8', annual.getCell('G23').result ?? 0);
    percentage(annual, 'C6', 'B6', '$B$5', otherExpenses, annualRevenue);
    percentage(annual, 'C8', 'B8', '$B$5', remuneration, annualRevenue);
    annual.mergeCells('A24:G24');
    annual.getCell('A24').value = 'Résultat = CA − rémunération − autres dépenses. — : CA nul. Le total des % est calculé sur le CA annuel.';
    annual.getCell('A24').font = { name: 'Arial', size: 10, italic: true, color: { argb: 'FF667085' } };
    annual.getRow(23).font = { name: 'Arial', size: 11, bold: true };
    finish(annual);
    book.views = [{ activeTab: 12 }];
    return book;
  }
  return { categories, months, cents, sum, summary, fetchAll, validateExpense, buildWorkbook };
});
