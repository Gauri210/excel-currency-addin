/*
 * Person 4: the side panel.
 * Connects the form and the Convert button to Person 2's code (converter.js)
 * and Person 3's code (writer.js), and shows "Loading…", errors and the result.
 */

/* global document, Office */

import {
  convert,
  fetchCurrencies,
  fetchRateHistory,
  shiftDate,
  todayISO,
  validateInput,
  ConversionError,
  EARLIEST_DATE,
} from "../currency/converter.js";
import { writeConversion, writeRateHistory } from "../excel/writer.js";

// Used if the currency list cannot be downloaded.
const FALLBACK_CURRENCIES = [
  "AUD", "BRL", "CAD", "CHF", "CNY", "CZK", "DKK", "EUR", "GBP", "HKD", "HUF", "IDR",
  "ILS", "INR", "ISK", "JPY", "KRW", "MXN", "MYR", "NOK", "NZD", "PHP", "PLN", "RON",
  "SEK", "SGD", "THB", "TRY", "USD", "ZAR",
].map((code) => ({ code, name: code }));

const DEFAULT_FROM = "USD";
const DEFAULT_TO = "INR";
const HISTORY_DAYS = 30;

const $ = (id) => document.getElementById(id);

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    $("sideload-msg").hidden = true;
    $("app-body").hidden = false;
    init();
  }
});

async function init() {
  const dateInput = $("date");
  dateInput.min = EARLIEST_DATE;
  dateInput.max = todayISO();

  $("convert-form").addEventListener("submit", onConvert);
  $("swap").addEventListener("click", swapCurrencies);
  $("amount").addEventListener("input", () => $("amount").classList.remove("invalid"));

  await loadCurrencies();
}

/* ------------------------------------------------------------------ */
/* Currency dropdowns                                                  */
/* ------------------------------------------------------------------ */

async function loadCurrencies() {
  let currencies;
  try {
    currencies = await fetchCurrencies();
  } catch (e) {
    currencies = FALLBACK_CURRENCIES;
  }
  fillSelect($("from"), currencies, DEFAULT_FROM);
  fillSelect($("to"), currencies, DEFAULT_TO);
}

function fillSelect(select, currencies, selected) {
  select.innerHTML = "";
  currencies.forEach(({ code, name }) => {
    const option = document.createElement("option");
    option.value = code;
    option.textContent = name && name !== code ? `${code} – ${name}` : code;
    if (code === selected) option.selected = true;
    select.appendChild(option);
  });
}

function swapCurrencies() {
  const from = $("from");
  const to = $("to");
  [from.value, to.value] = [to.value, from.value];
}

/* ------------------------------------------------------------------ */
/* Convert button                                                      */
/* ------------------------------------------------------------------ */

function readForm() {
  return {
    amount: $("amount").value,
    from: $("from").value,
    to: $("to").value,
    date: $("date").value, // "" when empty → latest rate
  };
}

async function onConvert(event) {
  event.preventDefault();
  hideResult();

  const input = readForm();

  // Check the input first so the user gets instant feedback.
  const problems = validateInput(input);
  if (problems.length > 0) {
    if (problems.some((p) => p.toLowerCase().includes("amount"))) $("amount").classList.add("invalid");
    showError(problems.join(" "));
    return;
  }

  setBusy(true);
  try {
    showLoading("Getting the exchange rate…");
    const conversion = await convert(input); // Person 2

    showLoading("Writing to the sheet…");
    await writeConversion(conversion); // Person 3

    let historyNote = "";
    if ($("show-history").checked) {
      showLoading("Drawing the rate chart…");
      historyNote = await addHistoryChart(conversion);
    }

    hideStatus();
    showResult(conversion, historyNote);
  } catch (error) {
    showError(friendlyMessage(error));
  } finally {
    setBusy(false);
  }
}

/**
 * The chart is extra: if it fails, keep the conversion and just say so.
 * @returns {Promise<string>} a short note for the result box
 */
async function addHistoryChart(conversion) {
  try {
    const end = conversion.rateDate;
    const start = shiftDate(end, -HISTORY_DAYS);
    const points = await fetchRateHistory(conversion.from, conversion.to, start, end); // Person 2
    await writeRateHistory(conversion.from, conversion.to, points); // Person 3
    return ` Chart added on the "Rate History" sheet.`;
  } catch (e) {
    return " (The rate chart could not be added.)";
  }
}

function friendlyMessage(error) {
  if (error instanceof ConversionError) return error.message;
  if (error && error.code === "InvalidOperationInCellEditMode") {
    return "Excel is busy editing a cell. Press Enter or Esc in the sheet, then try again.";
  }
  return `Something went wrong: ${(error && error.message) || error}`;
}

/* ------------------------------------------------------------------ */
/* Messages on screen                                                  */
/* ------------------------------------------------------------------ */

function setBusy(busy) {
  const button = $("convert");
  button.disabled = busy;
  button.textContent = busy ? "Loading…" : "Convert";
}

function showLoading(text) {
  const status = $("status");
  status.className = "status loading";
  status.textContent = text;
  status.hidden = false;
}

function showError(text) {
  const status = $("status");
  status.className = "status error";
  status.textContent = text;
  status.hidden = false;
}

function hideStatus() {
  $("status").hidden = true;
}

function hideResult() {
  $("result").hidden = true;
}

function formatMoney(value, currency) {
  try {
    return new Intl.NumberFormat(undefined, { style: "currency", currency }).format(value);
  } catch (e) {
    return `${value.toFixed(2)} ${currency}`;
  }
}

function formatDate(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, m - 1, d).toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" });
}

function showResult(c, note) {
  $("result-main").textContent = `${formatMoney(c.amount, c.from)} = ${formatMoney(c.result, c.to)}`;
  $("result-detail").textContent =
    `1 ${c.from} = ${c.rate} ${c.to}, rate of ${formatDate(c.rateDate)}. ` + `Added to the "Conversions" sheet.${note}`;
  $("result").hidden = false;
}
