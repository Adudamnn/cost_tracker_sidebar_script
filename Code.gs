4/*************** TALENTEASE COST TRACKER ***************/

const SALARY_SPREADSHEET_ID =
  '1do7d9c2GNwoBmr0HRpF1L8ccfEB9cNI1';

const SALARY_HEADER_ROW = 3;
const EMPLOYEE_NAME_COLUMN = 1; // Column A
const TOTAL_PAYABLE_COLUMN = 11; // Column K


/**************** MENU ****************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 Cost Tracker')
    .addItem('Open Cost Tracker', 'showSidebar')
    .addToUi();
}


/**************** SIDEBAR ****************/

function showSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile('Sidebar')
    .setTitle('TalentEase Cost Tracker');

  SpreadsheetApp.getUi().showSidebar(html);
}


/**************** INITIAL DATA ****************/

function getInitialData() {
  const salarySS = SpreadsheetApp.openById(SALARY_SPREADSHEET_ID);

  const months = salarySS
    .getSheets()
    .map(sheet => sheet.getName())
    .filter(name => /^\w+\s+\d{4}$/.test(name));

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let dashboard = ss.getSheetByName('Master_Dashboard');

  let clients = [];

  if (dashboard) {
    const lastRow = dashboard.getLastRow();

    if (lastRow >= 2) {
      clients = dashboard
        .getRange(2, 1, lastRow - 1, 1)
        .getDisplayValues()
        .flat()
        .map(x => x.trim())
        .filter(Boolean);
    }
  }

  return {
    months: months,
    clients: clients
  };
}


/**************** EMPLOYEES ****************/

function getEmployees(monthName) {
  const salarySS = SpreadsheetApp.openById(SALARY_SPREADSHEET_ID);
  const sheet = salarySS.getSheetByName(monthName);

  if (!sheet) {
    throw new Error('Salary sheet "' + monthName + '" was not found.');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow <= SALARY_HEADER_ROW) {
    return [];
  }

  const names = sheet
    .getRange(
      SALARY_HEADER_ROW + 1,
      EMPLOYEE_NAME_COLUMN,
      lastRow - SALARY_HEADER_ROW,
      1
    )
    .getDisplayValues()
    .flat()
    .map(x => x.trim())
    .filter(name => name && name.toLowerCase() !== 'total');

  return [...new Set(names)];
}


/**************** EMPLOYEE SALARY ****************/

function getEmployeeSalary(monthName, employeeName) {
  const salarySS = SpreadsheetApp.openById(SALARY_SPREADSHEET_ID);
  const sheet = salarySS.getSheetByName(monthName);

  if (!sheet) {
    throw new Error('Salary sheet "' + monthName + '" was not found.');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow <= SALARY_HEADER_ROW) {
    return 0;
  }

  const data = sheet
    .getRange(
      SALARY_HEADER_ROW + 1,
      1,
      lastRow - SALARY_HEADER_ROW,
      TOTAL_PAYABLE_COLUMN
    )
    .getDisplayValues();

  for (let i = 0; i < data.length; i++) {
    const employee = String(data[i][EMPLOYEE_NAME_COLUMN - 1]).trim();

    if (employee.toLowerCase() === employeeName.trim().toLowerCase()) {
      const salaryText = String(
        data[i][TOTAL_PAYABLE_COLUMN - 1]
      ).replace(/[₹,\s]/g, '');

      return Number(salaryText) || 0;
    }
  }

  throw new Error(
    'Employee "' + employeeName + '" was not found in ' + monthName + '.'
  );
}


/**************** CLIENT BUDGET ****************/

function getClientBudget(clientName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('Master_Dashboard');

  if (!dashboard) {
    throw new Error('Master_Dashboard sheet was not found.');
  }

  const lastRow = dashboard.getLastRow();

  if (lastRow < 2) {
    return 0;
  }

  const data = dashboard
    .getRange(2, 1, lastRow - 1, 2)
    .getDisplayValues();

  for (let i = 0; i < data.length; i++) {
    if (
      String(data[i][0]).trim().toLowerCase() ===
      clientName.trim().toLowerCase()
    ) {
      return parseMoney(data[i][1]);
    }
  }

  return 0;
}


/**************** SET CLIENT BUDGET ****************/

function setClientBudget(clientName, budget) {
  if (!clientName) {
    throw new Error('Please select a client.');
  }

  budget = Number(budget);

  if (isNaN(budget) || budget < 0) {
    throw new Error('Please enter a valid budget.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('Master_Dashboard');

  if (!dashboard) {
    throw new Error('Master_Dashboard sheet was not found.');
  }

  const lastRow = dashboard.getLastRow();

  if (lastRow < 2) {
    throw new Error('No clients were found in Master_Dashboard.');
  }

  const clients = dashboard
    .getRange(2, 1, lastRow - 1, 1)
    .getDisplayValues()
    .flat();

  for (let i = 0; i < clients.length; i++) {
    if (
      String(clients[i]).trim().toLowerCase() ===
      clientName.trim().toLowerCase()
    ) {
      dashboard.getRange(i + 2, 2).setValue(budget);
      return true;
    }
  }

  throw new Error('Client "' + clientName + '" was not found.');
}


/**************** CREATE ENTRIES SHEET ****************/

function getEntriesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('Cost_Entries');

  if (!sheet) {
    sheet = ss.insertSheet('Cost_Entries');

    sheet.getRange(1, 1, 1, 8).setValues([[
      'Timestamp',
      'Month',
      'Employee',
      'Total Payable Salary',
      'Client',
      'Category',
      'Allocation %',
      'Allocated Cost'
    ]]);

    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }

  return sheet;
}


/**************** SUBMIT ENTRY ****************/

function submitEntry(entry) {
  if (!entry) {
    throw new Error('No entry received.');
  }

  if (!entry.month) {
    throw new Error('Please select a month.');
  }

  if (!entry.employee) {
    throw new Error('Please select an employee.');
  }

  if (!entry.client) {
    throw new Error('Please select a client.');
  }

  if (!entry.category) {
    throw new Error('Please select a category.');
  }

  const allocation = Number(entry.allocation);

  if (
    isNaN(allocation) ||
    allocation <= 0 ||
    allocation > 100
  ) {
    throw new Error('Allocation must be between 1% and 100%.');
  }

  const salary = getEmployeeSalary(
    entry.month,
    entry.employee
  );

  const allocatedCost = salary * allocation / 100;

  const sheet = getEntriesSheet();

  sheet.appendRow([
    new Date(),
    entry.month,
    entry.employee,
    salary,
    entry.client,
    entry.category,
    allocation / 100,
    allocatedCost
  ]);

  return getClientSummary(
    entry.month,
    entry.client
  );
}

/**************** CLIENT SUMMARY ****************/

function getClientSummary(monthName, clientName) {

  const budget = getClientBudget(clientName);
  const sheet = getEntriesSheet();

  const lastRow = sheet.getLastRow();

  let monthlySpend = 0;

  if (lastRow >= 2) {

    const data = sheet
      .getRange(2, 1, lastRow - 1, 8)
      .getDisplayValues();

    const selectedMonth = normalizeText(monthName);
    const selectedClient = normalizeText(clientName);

    data.forEach(function(row) {

      const rowMonth = normalizeText(row[1]);
      const rowClient = normalizeText(row[4]);

      if (
        rowMonth === selectedMonth &&
        rowClient === selectedClient
      ) {

        const allocatedCost = parseMoney(row[7]);

        monthlySpend += allocatedCost;
      }
    });
  }

  const remaining = budget - monthlySpend;

  const percentageUsed =
    budget > 0
      ? (monthlySpend / budget) * 100
      : 0;

  return {
    budget: budget,
    monthlySpend: monthlySpend,
    percentageUsed: percentageUsed,
    remaining: remaining,
    overBudget: remaining < 0
  };
}

/**************** TEXT HELPER ****************/

function normalizeText(value) {

  return String(value || '')
    .replace(/\u00A0/g, ' ')
    .trim()
    .toLowerCase();
}

/**************** MONEY HELPER ****************/

function parseMoney(value) {
  if (typeof value === 'number') {
    return value;
  }

  return Number(
    String(value)
      .replace(/[₹,\s]/g, '')
      .replace(/[^\d.-]/g, '')
  ) || 0;
}