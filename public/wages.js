const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];

const el = {
  tabs: $$(".wage-tab"),
  employeesSheet: $("#employeesSheet"),
  contractorsSheet: $("#contractorsSheet"),
  employeeWageForm: $("#employeeWageForm"),
  employeeWageId: $("#employeeWageId"),
  employeeName: $("#employeeNameInput"),
  employmentType: $("#employmentTypeInput"),
  payPeriod: $("#payPeriodInput"),
  employeeAmount: $("#employeeAmountInput"),
  employeeCurrency: $("#employeeCurrencyInput"),
  employeeNotes: $("#employeeNotesInput"),
  employeeWageMessage: $("#employeeWageMessage"),
  employeeWageTotal: $("#employeeWageTotal"),
  employeeWageList: $("#employeeWageList"),
  employeeWageTemplate: $("#employeeWageTemplate"),
  newEmployeeWageButton: $("#newEmployeeWageButton"),
  wageList: $("#wageList"),
  wageTemplate: $("#wageTemplate"),
  wageMessage: $("#wageMessage"),
  wageDraftCount: $("#wageDraftCount"),
  wagePaidCount: $("#wagePaidCount"),
  wageTotalValue: $("#wageTotalValue"),
  syncWagesButton: $("#syncWagesButton"),
  refreshWagesButton: $("#refreshWagesButton"),
  logoutButton: $("#logoutButton")
};

const state = {
  activeSheet: "employees",
  employeeWages: [],
  contractorWages: []
};

const moneyFormatters = new Map();

function formatMoney(value, currency = "AUD") {
  const key = currency || "AUD";
  if (!moneyFormatters.has(key)) {
    moneyFormatters.set(key, new Intl.NumberFormat("en-AU", {
      style: "currency",
      currency: key,
      currencyDisplay: key === "THB" ? "narrowSymbol" : "symbol"
    }));
  }
  return moneyFormatters.get(key).format(Number(value || 0));
}

function formatDate(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return new Intl.DateTimeFormat("en-AU", { day: "numeric", month: "short", year: "numeric" }).format(date);
}

function labelize(value) {
  return String(value || "")
    .replace(/_/g, " ")
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function parseEmails(value) {
  return String(value || "")
    .split(/[\s,;]+/)
    .map((email) => email.trim())
    .filter(Boolean);
}

function isEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function uniqueEmails(values) {
  const seen = new Set();
  const emails = [];
  for (const value of values) {
    const key = String(value || "").trim().toLowerCase();
    if (!key || seen.has(key)) continue;
    seen.add(key);
    emails.push(String(value).trim());
  }
  return emails;
}

function setMessage(target, message, type = "") {
  target.textContent = message || "";
  target.className = type;
}

async function apiFetch(url, options = {}) {
  const response = await fetch(url, {
    credentials: "include",
    headers: {
      Accept: "application/json",
      ...(options.body ? { "Content-Type": "application/json" } : {}),
      ...(options.headers || {})
    },
    ...options
  });
  const data = await response.json().catch(() => ({}));
  if (response.status === 401) {
    window.location.href = "/login";
    throw new Error("Log in first.");
  }
  return { response, data };
}

function jsonRequest(method, body = null) {
  return {
    method,
    ...(body ? { body: JSON.stringify(body) } : {})
  };
}

function setActiveSheet(sheet) {
  state.activeSheet = sheet;
  el.tabs.forEach((tab) => tab.classList.toggle("active", tab.dataset.sheet === sheet));
  el.employeesSheet.hidden = sheet !== "employees";
  el.contractorsSheet.hidden = sheet !== "contractors";
}

function employeePayload() {
  return {
    id: el.employeeWageId.value,
    employeeName: el.employeeName.value.trim(),
    employmentType: el.employmentType.value,
    payPeriod: el.payPeriod.value,
    amount: Number(el.employeeAmount.value),
    currency: el.employeeCurrency.value.trim().toUpperCase(),
    notes: el.employeeNotes.value.trim()
  };
}

function resetEmployeeForm(message = "") {
  el.employeeWageId.value = "";
  el.employeeName.value = "Faye";
  el.employmentType.value = "full_time";
  el.payPeriod.value = "monthly";
  el.employeeAmount.value = "15000";
  el.employeeCurrency.value = "THB";
  el.employeeNotes.value = "";
  setMessage(el.employeeWageMessage, message);
}

function fillEmployeeForm(wage) {
  el.employeeWageId.value = wage.id || "";
  el.employeeName.value = wage.employeeName || "";
  el.employmentType.value = wage.employmentType || "full_time";
  el.payPeriod.value = wage.payPeriod || "monthly";
  el.employeeAmount.value = Number(wage.amount || 0);
  el.employeeCurrency.value = wage.currency || "THB";
  el.employeeNotes.value = wage.notes || "";
  setMessage(el.employeeWageMessage, `Editing ${wage.employeeName}.`);
  el.employeeWageForm.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderEmployeeWages() {
  el.employeeWageList.innerHTML = "";
  const wages = [...state.employeeWages].sort((a, b) => a.employeeName.localeCompare(b.employeeName));
  const activeTotal = wages
    .filter((wage) => wage.status !== "void")
    .reduce((sum, wage) => sum + Number(wage.amount || wage.total || 0), 0);
  const currency = wages.find((wage) => wage.status !== "void")?.currency || "THB";
  el.employeeWageTotal.textContent = `${formatMoney(activeTotal, currency)} ${currency}`;

  if (!wages.length) {
    const empty = document.createElement("p");
    empty.className = "empty-state";
    empty.textContent = "No employee wages yet.";
    el.employeeWageList.append(empty);
    return;
  }

  for (const wage of wages) {
    const item = el.employeeWageTemplate.content.firstElementChild.cloneNode(true);
    item.classList.toggle("paid", wage.status === "paid");
    item.classList.toggle("void", wage.status === "void");
    item.querySelector("h3").textContent = `${wage.wageNumber || "Employee"} - ${wage.employeeName}`;
    const status = item.querySelector(".invoice-status");
    status.textContent = labelize(wage.status || "draft");
    status.classList.toggle("paid", wage.status === "paid");
    status.classList.toggle("void", wage.status === "void");
    item.querySelector(".employee-wage-type").textContent = `${labelize(wage.employmentType)} · ${labelize(wage.payPeriod)} · Issued ${formatDate(wage.issuedAt)}`;
    item.querySelector(".employee-wage-notes").textContent = wage.notes || "No notes";
    item.querySelector(".invoice-total strong").textContent = formatMoney(wage.amount || wage.total, wage.currency || "THB");
    item.querySelector(".invoice-total small").textContent = `${wage.currency || "THB"} ${labelize(wage.payPeriod)}`;
    item.querySelector(".print-employee-payslip-button").addEventListener("click", () => printEmployeePayslip(wage.id));
    item.querySelector(".edit-employee-wage-button").addEventListener("click", () => fillEmployeeForm(wage));
    item.querySelector(".paid-employee-wage-button").hidden = wage.status !== "draft";
    item.querySelector(".paid-employee-wage-button").addEventListener("click", () => updateEmployeeWageStatus(wage.id, "paid"));
    item.querySelector(".draft-employee-wage-button").hidden = wage.status === "draft";
    item.querySelector(".draft-employee-wage-button").addEventListener("click", () => updateEmployeeWageStatus(wage.id, "draft"));
    item.querySelector(".void-employee-wage-button").hidden = wage.status === "void";
    item.querySelector(".void-employee-wage-button").addEventListener("click", () => updateEmployeeWageStatus(wage.id, "void"));
    item.querySelector(".delete-employee-wage-button").addEventListener("click", () => deleteEmployeeWage(wage.id));
    el.employeeWageList.append(item);
  }

  if (window.lucide) window.lucide.createIcons();
}

function invoiceStatusLabel(status) {
  return status === "paid" ? "Paid" : status === "void" ? "Void" : "Draft";
}

function wageRecipientEmails(wage) {
  return uniqueEmails(parseEmails(wage.photographerEmail || "")).filter(isEmail);
}

function updateContractorStats() {
  const drafts = state.contractorWages.filter((wage) => wage.status === "draft");
  const paid = state.contractorWages.filter((wage) => wage.status === "paid");
  el.wageDraftCount.textContent = drafts.length;
  el.wagePaidCount.textContent = paid.length;
  el.wageTotalValue.textContent = formatMoney(drafts.reduce((sum, wage) => sum + Number(wage.total || 0), 0), "AUD");
}

function renderContractorWages() {
  el.wageList.innerHTML = "";
  const wages = [...state.contractorWages].sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
  updateContractorStats();

  if (!wages.length) {
    const empty = document.createElement("p");
    empty.className = "empty-state";
    empty.textContent = "No contractor wages yet. Sync bookings to create photographer proformas.";
    el.wageList.append(empty);
    return;
  }

  for (const wage of wages) {
    const item = el.wageTemplate.content.firstElementChild.cloneNode(true);
    item.classList.toggle("paid", wage.status === "paid");
    item.classList.toggle("void", wage.status === "void");
    item.querySelector("h3").textContent = `${wage.wageNumber} - ${wage.propertyAddress || "Photographer wage"}`;
    const status = item.querySelector(".invoice-status");
    status.textContent = invoiceStatusLabel(wage.status);
    status.classList.toggle("paid", wage.status === "paid");
    status.classList.toggle("void", wage.status === "void");
    item.querySelector(".wage-photographer").textContent = [wage.photographerName, wage.photographerEmail].filter(Boolean).join(" - ") || "No photographer email";
    item.querySelector(".wage-booking").textContent = `Issued ${formatDate(wage.issuedAt)} · ${wage.clientName || "Booking"}`;
    item.querySelector(".wage-services").textContent = [
      (wage.items || []).map((line) => `${line.name} ${formatMoney(line.amount, wage.currency || "AUD")}`).join(" + "),
      wage.photographerGstIncluded ? `GST included ${formatMoney(wage.gstAmount || 0, wage.currency || "AUD")}` : "No GST"
    ].filter(Boolean).join(" · ");
    item.querySelector(".wage-sent").textContent = wage.sentAt
      ? `Sent ${formatDate(wage.sentAt)}${wage.sentTo?.length ? ` to ${wage.sentTo.join(", ")}` : ""}`
      : "Not sent";
    item.querySelector(".invoice-total strong").textContent = formatMoney(wage.total, wage.currency || "AUD");
    item.querySelector(".invoice-total small").textContent = wage.currency || "AUD";
    item.querySelector(".print-wage-button").addEventListener("click", () => printContractorWage(wage.id));
    item.querySelector(".send-wage-button").hidden = wage.status === "void";
    item.querySelector(".send-wage-button").addEventListener("click", () => sendContractorWage(wage.id));
    item.querySelector(".paid-wage-button").hidden = wage.status !== "draft";
    item.querySelector(".paid-wage-button").addEventListener("click", () => updateContractorWageStatus(wage.id, "paid"));
    item.querySelector(".void-wage-button").hidden = wage.status === "void";
    item.querySelector(".void-wage-button").addEventListener("click", () => updateContractorWageStatus(wage.id, "void"));
    el.wageList.append(item);
  }
}

async function loadEmployeeWages() {
  const { response, data } = await apiFetch("/api/employee-wages");
  if (!response.ok) {
    setMessage(el.employeeWageMessage, (data.errors || ["Could not load employee wages."]).join(" "), "error");
    return;
  }
  state.employeeWages = data.employeeWages || [];
  renderEmployeeWages();
}

async function saveEmployeeWage(event) {
  event.preventDefault();
  setMessage(el.employeeWageMessage, "Saving employee wage...");
  const { response, data } = await apiFetch("/api/employee-wages", jsonRequest("POST", employeePayload()));
  if (!response.ok) {
    setMessage(el.employeeWageMessage, (data.errors || ["Could not save employee wage."]).join(" "), "error");
    return;
  }
  state.employeeWages = data.employeeWages || [];
  renderEmployeeWages();
  resetEmployeeForm("Employee wage saved.");
  setMessage(el.employeeWageMessage, "Employee wage saved.", "success");
}

async function printEmployeePayslip(id) {
  const wage = state.employeeWages.find((item) => item.id === id);
  if (!wage) return;
  setMessage(el.employeeWageMessage, `Preparing ${wage.employeeName} payslip...`);
  const response = await fetch(`/api/employee-wages/${encodeURIComponent(id)}/pdf`, { credentials: "include" });
  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    setMessage(el.employeeWageMessage, (data.errors || ["Could not create payslip PDF."]).join(" "), "error");
    return;
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${[wage.wageNumber, wage.employeeName, "Pay Slip"].filter(Boolean).join(" - ").replace(/[\\/:*?"<>|]/g, " - ")}.pdf`;
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1500);
  setMessage(el.employeeWageMessage, `${wage.employeeName} payslip downloaded.`, "success");
}

async function updateEmployeeWageStatus(id, status) {
  const { response, data } = await apiFetch(`/api/employee-wages/${encodeURIComponent(id)}/${status}`, { method: "POST" });
  if (!response.ok) {
    setMessage(el.employeeWageMessage, (data.errors || ["Could not update employee wage."]).join(" "), "error");
    return;
  }
  state.employeeWages = data.employeeWages || [];
  renderEmployeeWages();
  setMessage(el.employeeWageMessage, "Employee wage updated.", "success");
}

async function deleteEmployeeWage(id) {
  const wage = state.employeeWages.find((item) => item.id === id);
  if (!wage || !window.confirm(`Delete ${wage.employeeName}'s wage row?`)) return;
  const { response, data } = await apiFetch(`/api/employee-wages/${encodeURIComponent(id)}`, { method: "DELETE" });
  if (!response.ok) {
    setMessage(el.employeeWageMessage, (data.errors || ["Could not delete employee wage."]).join(" "), "error");
    return;
  }
  state.employeeWages = data.employeeWages || [];
  renderEmployeeWages();
  resetEmployeeForm("Employee wage deleted.");
}

async function loadContractorWages() {
  const { response, data } = await apiFetch("/api/wages");
  if (!response.ok) {
    setMessage(el.wageMessage, (data.errors || ["Could not load contractor wages."]).join(" "), "error");
    return;
  }
  state.contractorWages = data.wages || [];
  renderContractorWages();
}

async function syncContractorWages() {
  el.syncWagesButton.disabled = true;
  setMessage(el.wageMessage, "Creating missing contractor wages...");
  try {
    const { response, data } = await apiFetch("/api/wages/sync", { method: "POST" });
    if (!response.ok) {
      setMessage(el.wageMessage, (data.errors || ["Could not sync contractor wages."]).join(" "), "error");
      return;
    }
    state.contractorWages = data.wages || [];
    renderContractorWages();
    setMessage(el.wageMessage, `Contractor wages synced: ${data.created || 0} created, ${data.updated || 0} updated.`, "success");
  } finally {
    el.syncWagesButton.disabled = false;
  }
}

async function printContractorWage(id) {
  const wage = state.contractorWages.find((item) => item.id === id);
  if (!wage) return;
  setMessage(el.wageMessage, `Preparing ${wage.wageNumber} PDF...`);
  const response = await fetch(`/api/wages/${encodeURIComponent(id)}/pdf`, { credentials: "include" });
  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    setMessage(el.wageMessage, (data.errors || ["Could not create wage PDF."]).join(" "), "error");
    return;
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${[wage.wageNumber, wage.propertyAddress, "Photographer Proforma"].filter(Boolean).join(" - ").replace(/[\\/:*?"<>|]/g, " - ")}.pdf`;
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1500);
  setMessage(el.wageMessage, `${wage.wageNumber} PDF downloaded.`, "success");
}

async function sendContractorWage(id) {
  const wage = state.contractorWages.find((item) => item.id === id);
  if (!wage) return;
  const recipients = wageRecipientEmails(wage);
  if (!recipients.length) {
    setMessage(el.wageMessage, "Add a photographer email before sending this proforma.", "error");
    return;
  }
  if (!window.confirm(`Send ${wage.wageNumber} to ${recipients.join(", ")}?`)) return;
  setMessage(el.wageMessage, `Sending ${wage.wageNumber}...`);
  const { response, data } = await apiFetch(`/api/wages/${encodeURIComponent(id)}/send`, jsonRequest("POST", { to: recipients }));
  if (!response.ok) {
    setMessage(el.wageMessage, (data.errors || ["Could not send proforma."]).join(" "), "error");
    return;
  }
  state.contractorWages = data.wages || state.contractorWages.map((item) => item.id === id ? data.wage : item);
  renderContractorWages();
  setMessage(el.wageMessage, data.message || "Proforma sent.", "success");
}

async function updateContractorWageStatus(id, status) {
  const { response, data } = await apiFetch(`/api/wages/${encodeURIComponent(id)}/${status}`, { method: "POST" });
  if (!response.ok) {
    setMessage(el.wageMessage, (data.errors || ["Could not update contractor wage."]).join(" "), "error");
    return;
  }
  state.contractorWages = data.wages || [];
  renderContractorWages();
  setMessage(el.wageMessage, status === "paid" ? "Contractor wage marked paid." : "Contractor wage voided.", "success");
}

async function logout() {
  await fetch("/api/logout", { method: "POST", credentials: "include" });
  window.location.href = "/login";
}

function wireEvents() {
  el.tabs.forEach((tab) => tab.addEventListener("click", () => setActiveSheet(tab.dataset.sheet)));
  el.employeeWageForm.addEventListener("submit", saveEmployeeWage);
  el.newEmployeeWageButton.addEventListener("click", () => resetEmployeeForm());
  el.refreshWagesButton.addEventListener("click", loadContractorWages);
  el.syncWagesButton.addEventListener("click", syncContractorWages);
  el.logoutButton.addEventListener("click", logout);
}

wireEvents();
setActiveSheet("employees");
await Promise.all([loadEmployeeWages(), loadContractorWages()]);

if (window.lucide) window.lucide.createIcons();
window.addEventListener("load", () => {
  if (window.lucide) window.lucide.createIcons();
});
