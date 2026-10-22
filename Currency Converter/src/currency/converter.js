/*
 * Person 2: get the exchange rate and do the maths.
 *
 * Everything here is plain JavaScript with no Excel code, so it can be tested
 * on its own with `npm test`.
 *
 * Data source: Frankfurter API (https://frankfurter.dev), free, no API key.
 * Rates are published by the European Central Bank on working days only. If
 * you ask for a weekend or holiday, the API returns the rate of the previous
 * working day, and the "date" field in the answer tells you which day it was.
 */

export const API_BASE = "https://api.frankfurter.dev/v1";

// The first day the API has data for.
export const EARLIEST_DATE = "1999-01-04";

/* ------------------------------------------------------------------ */
/* Errors                                                              */
/* ------------------------------------------------------------------ */

/** An error whose message is safe to show to the user as it is. */
export class ConversionError extends Error {
  constructor(message) {
    super(message);
    this.name = "ConversionError";
  }
}

/* ------------------------------------------------------------------ */
/* Input checks                                                        */
/* ------------------------------------------------------------------ */

const CURRENCY_CODE = /^[A-Z]{3}$/;
const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

/** Today's date as YYYY-MM-DD in the user's own time zone. */
export function todayISO(now = new Date()) {
  const pad = (n) => String(n).padStart(2, "0");
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
}

/** True if text is a real calendar date written as YYYY-MM-DD. */
export function isValidISODate(text) {
  if (typeof text !== "string" || !ISO_DATE.test(text)) return false;
  const [y, m, d] = text.split("-").map(Number);
  const date = new Date(Date.UTC(y, m - 1, d));
  return date.getUTCFullYear() === y && date.getUTCMonth() === m - 1 && date.getUTCDate() === d;
}

/**
 * Turn whatever the user typed into a number.
 * Accepts "1000", "1,000.50", " 25 ". Returns NaN for anything else.
 */
export function parseAmount(raw) {
  if (typeof raw === "number") return raw;
  if (typeof raw !== "string") return NaN;
  const cleaned = raw.trim().replace(/,/g, "");
  if (cleaned === "" || !/^[-+]?(\d+\.?\d*|\.\d+)$/.test(cleaned)) return NaN;
  return Number(cleaned);
}

/**
 * Check the form. Returns a list of problems (empty list = all good).
 *
 * @param {{amount: string|number, from: string, to: string, date?: string}} input
 * @param {Date} [now] only passed in by tests
 * @returns {string[]}
 */
export function validateInput(input, now = new Date()) {
  const errors = [];
  const { amount, from, to, date } = input || {};

  if (amount === undefined || amount === null || String(amount).trim() === "") {
    errors.push("Please enter an amount.");
  } else {
    const value = parseAmount(amount);
    if (!Number.isFinite(value)) {
      errors.push("The amount must be a number.");
    } else if (value < 0) {
      errors.push("The amount cannot be negative.");
    } else if (value === 0) {
      errors.push("The amount must be greater than zero.");
    }
  }

  if (!from) errors.push("Please choose the currency to convert from.");
  else if (!CURRENCY_CODE.test(from)) errors.push(`"${from}" is not a valid currency code.`);

  if (!to) errors.push("Please choose the currency to convert to.");
  else if (!CURRENCY_CODE.test(to)) errors.push(`"${to}" is not a valid currency code.`);

  if (from && to && from === to) {
    errors.push("Choose two different currencies.");
  }

  if (date) {
    if (!isValidISODate(date)) {
      errors.push("The date is not valid.");
    } else if (date > todayISO(now)) {
      errors.push("The date cannot be in the future.");
    } else if (date < EARLIEST_DATE) {
      errors.push("Rates are only available from 4 January 1999.");
    }
  }

  return errors;
}

/* ------------------------------------------------------------------ */
/* The maths                                                           */
/* ------------------------------------------------------------------ */

/** Round to a number of decimal places without floating-point surprises. */
export function roundTo(value, decimals = 2) {
  const factor = 10 ** decimals;
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

/**
 * result = amount × rate, rounded to 2 decimal places.
 *
 * @param {number} amount
 * @param {number} rate
 * @returns {number}
 */
export function calculate(amount, rate) {
  if (!Number.isFinite(amount) || amount < 0) {
    throw new ConversionError("The amount must be a positive number.");
  }
  if (!Number.isFinite(rate) || rate <= 0) {
    throw new ConversionError("The exchange rate is not valid.");
  }
  return roundTo(amount * rate, 2);
}

/* ------------------------------------------------------------------ */
/* Talking to the API                                                  */
/* ------------------------------------------------------------------ */

/**
 * Fetch JSON and turn every kind of failure into a friendly ConversionError.
 * `fetchFn` lets the tests pass a fake fetch instead of going online.
 */
async function getJSON(url, fetchFn = globalThis.fetch) {
  let response;
  try {
    response = await fetchFn(url);
  } catch (e) {
    throw new ConversionError("Could not reach the exchange-rate service. Check your internet connection.");
  }

  if (response.status === 404 || response.status === 422) {
    throw new ConversionError("No rate is available for that currency or date.");
  }
  if (!response.ok) {
    throw new ConversionError(`The exchange-rate service returned an error (${response.status}). Try again later.`);
  }

  try {
    return await response.json();
  } catch (e) {
    throw new ConversionError("The exchange-rate service sent back an unreadable answer.");
  }
}

/**
 * Get the rate for 1 unit of `from` in `to`.
 * Leave `date` empty for the latest rate.
 *
 * @returns {Promise<{rate: number, date: string}>} date = the day the rate is from
 */
export async function fetchRate(from, to, date, fetchFn) {
  const path = date ? date : "latest";
  const url = `${API_BASE}/${path}?base=${encodeURIComponent(from)}&symbols=${encodeURIComponent(to)}`;
  const data = await getJSON(url, fetchFn);

  const rate = data && data.rates ? data.rates[to] : undefined;
  if (typeof rate !== "number" || !(rate > 0)) {
    throw new ConversionError(`No rate was found for ${from} to ${to}.`);
  }
  return { rate, date: data.date };
}

/**
 * Get the rate for every working day between two dates (for the chart).
 *
 * @returns {Promise<Array<{date: string, rate: number}>>} oldest first
 */
export async function fetchRateHistory(from, to, startDate, endDate, fetchFn) {
  const url =
    `${API_BASE}/${startDate}..${endDate}` +
    `?base=${encodeURIComponent(from)}&symbols=${encodeURIComponent(to)}`;
  const data = await getJSON(url, fetchFn);

  if (!data || !data.rates) {
    throw new ConversionError("No rate history was returned.");
  }
  return Object.keys(data.rates)
    .sort()
    .map((day) => ({ date: day, rate: data.rates[day][to] }))
    .filter((point) => typeof point.rate === "number");
}

/**
 * Get the list of currencies the API supports.
 *
 * @returns {Promise<Array<{code: string, name: string}>>} sorted by code
 */
export async function fetchCurrencies(fetchFn) {
  const data = await getJSON(`${API_BASE}/currencies`, fetchFn);
  return Object.keys(data)
    .sort()
    .map((code) => ({ code, name: data[code] }));
}

/** The date `days` days before `isoDate`, as YYYY-MM-DD. */
export function shiftDate(isoDate, days) {
  const [y, m, d] = isoDate.split("-").map(Number);
  const date = new Date(Date.UTC(y, m - 1, d));
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

/* ------------------------------------------------------------------ */
/* Everything together                                                 */
/* ------------------------------------------------------------------ */

/**
 * Check the input, get the rate, and calculate the result.
 * This is the one function the task pane calls.
 *
 * @param {{amount: string|number, from: string, to: string, date?: string}} input
 * @returns {Promise<{amount: number, from: string, to: string, rate: number, rateDate: string, result: number}>}
 * @throws {ConversionError} with a message that can be shown to the user
 */
export async function convert(input, fetchFn) {
  const errors = validateInput(input);
  if (errors.length > 0) {
    throw new ConversionError(errors.join(" "));
  }

  const amount = parseAmount(input.amount);
  const { rate, date } = await fetchRate(input.from, input.to, input.date, fetchFn);

  return {
    amount,
    from: input.from,
    to: input.to,
    rate,
    rateDate: date,
    result: calculate(amount, rate),
  };
}
