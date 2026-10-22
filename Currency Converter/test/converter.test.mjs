/*
 * Person 2: tests that prove the conversion is correct.
 * Run with: npm test   (uses Node's built-in test runner, Node 20 or newer)
 *
 * The API is never called for real here. A fake `fetch` returns a fixed answer,
 * so the tests are fast and always give the same result.
 */
import { test, describe } from "node:test";
import assert from "node:assert/strict";
import {
  calculate,
  convert,
  ConversionError,
  fetchRate,
  fetchRateHistory,
  fetchCurrencies,
  isValidISODate,
  parseAmount,
  roundTo,
  shiftDate,
  validateInput,
} from "../src/currency/converter.js";

/** A fake fetch that remembers the URL it was asked for. */
function fakeFetch(body, status = 200) {
  const fn = async (url) => {
    fn.lastUrl = url;
    return {
      ok: status >= 200 && status < 300,
      status,
      json: async () => body,
    };
  };
  return fn;
}

const FIXED_NOW = new Date(2026, 9, 22); // 22 Oct 2026

describe("calculate (amount × rate)", () => {
  test("multiplies amount by rate", () => {
    assert.equal(calculate(100, 83.25), 8325);
  });

  test("rounds to 2 decimal places", () => {
    assert.equal(calculate(10, 0.012345), 0.12);
    assert.equal(calculate(1, 1.005), 1.01); // the classic floating-point trap
  });

  test("handles decimals in the amount", () => {
    assert.equal(calculate(19.99, 1.1), 21.99);
  });

  test("an amount of 1 returns the rate itself", () => {
    assert.equal(calculate(1, 0.9214), 0.92);
  });

  test("rejects a negative amount", () => {
    assert.throws(() => calculate(-5, 1.2), ConversionError);
  });

  test("rejects a zero or missing rate", () => {
    assert.throws(() => calculate(5, 0), ConversionError);
    assert.throws(() => calculate(5, NaN), ConversionError);
  });
});

describe("roundTo", () => {
  test("rounds half up", () => {
    assert.equal(roundTo(2.675, 2), 2.68);
    assert.equal(roundTo(1.23456, 4), 1.2346);
  });
});

describe("parseAmount", () => {
  test("reads plain and comma-separated numbers", () => {
    assert.equal(parseAmount("1000"), 1000);
    assert.equal(parseAmount("1,000.50"), 1000.5);
    assert.equal(parseAmount("  25 "), 25);
    assert.equal(parseAmount(".5"), 0.5);
  });

  test("returns NaN for text", () => {
    assert.ok(Number.isNaN(parseAmount("abc")));
    assert.ok(Number.isNaN(parseAmount("12abc")));
    assert.ok(Number.isNaN(parseAmount("")));
  });
});

describe("validateInput", () => {
  const good = { amount: "100", from: "USD", to: "INR", date: "2024-03-15" };

  test("accepts good input", () => {
    assert.deepEqual(validateInput(good, FIXED_NOW), []);
  });

  test("date is optional", () => {
    assert.deepEqual(validateInput({ ...good, date: "" }, FIXED_NOW), []);
  });

  test("empty amount", () => {
    assert.deepEqual(validateInput({ ...good, amount: "" }, FIXED_NOW), ["Please enter an amount."]);
    assert.deepEqual(validateInput({ ...good, amount: "   " }, FIXED_NOW), ["Please enter an amount."]);
  });

  test("negative amount", () => {
    assert.deepEqual(validateInput({ ...good, amount: "-10" }, FIXED_NOW), ["The amount cannot be negative."]);
  });

  test("zero amount", () => {
    assert.deepEqual(validateInput({ ...good, amount: "0" }, FIXED_NOW), ["The amount must be greater than zero."]);
  });

  test("amount that is not a number", () => {
    assert.deepEqual(validateInput({ ...good, amount: "ten" }, FIXED_NOW), ["The amount must be a number."]);
  });

  test("same currency twice", () => {
    assert.deepEqual(validateInput({ ...good, to: "USD" }, FIXED_NOW), ["Choose two different currencies."]);
  });

  test("missing currencies", () => {
    const errors = validateInput({ ...good, from: "", to: "" }, FIXED_NOW);
    assert.equal(errors.length, 2);
  });

  test("future date", () => {
    assert.deepEqual(validateInput({ ...good, date: "2026-10-23" }, FIXED_NOW), ["The date cannot be in the future."]);
  });

  test("date before the data starts", () => {
    assert.deepEqual(validateInput({ ...good, date: "1998-12-31" }, FIXED_NOW), [
      "Rates are only available from 4 January 1999.",
    ]);
  });

  test("impossible date", () => {
    assert.deepEqual(validateInput({ ...good, date: "2023-02-30" }, FIXED_NOW), ["The date is not valid."]);
  });
});

describe("dates", () => {
  test("isValidISODate", () => {
    assert.ok(isValidISODate("2024-02-29")); // leap year
    assert.ok(!isValidISODate("2023-02-29"));
    assert.ok(!isValidISODate("15/03/2024"));
  });

  test("shiftDate goes across month and year ends", () => {
    assert.equal(shiftDate("2024-03-01", -1), "2024-02-29");
    assert.equal(shiftDate("2025-01-01", -1), "2024-12-31");
    assert.equal(shiftDate("2024-03-15", -30), "2024-02-14");
  });
});

describe("fetchRate", () => {
  test("asks for the latest rate when there is no date", async () => {
    const fetch = fakeFetch({ amount: 1, base: "USD", date: "2026-10-21", rates: { INR: 84.12 } });
    const result = await fetchRate("USD", "INR", "", fetch);
    assert.equal(fetch.lastUrl, "https://api.frankfurter.dev/v1/latest?base=USD&symbols=INR");
    assert.deepEqual(result, { rate: 84.12, date: "2026-10-21" });
  });

  test("asks for a past date when one is given", async () => {
    const fetch = fakeFetch({ amount: 1, base: "EUR", date: "2024-03-15", rates: { GBP: 0.8545 } });
    await fetchRate("EUR", "GBP", "2024-03-15", fetch);
    assert.equal(fetch.lastUrl, "https://api.frankfurter.dev/v1/2024-03-15?base=EUR&symbols=GBP");
  });

  test("returns the date the rate is really from (weekend → Friday)", async () => {
    const fetch = fakeFetch({ amount: 1, base: "EUR", date: "2024-03-15", rates: { GBP: 0.8545 } });
    const result = await fetchRate("EUR", "GBP", "2024-03-17", fetch);
    assert.equal(result.date, "2024-03-15");
  });

  test("error when the currency is missing in the answer", async () => {
    const fetch = fakeFetch({ amount: 1, base: "USD", date: "2026-10-21", rates: {} });
    await assert.rejects(fetchRate("USD", "INR", "", fetch), ConversionError);
  });

  test("error for an unknown currency (404/422)", async () => {
    const fetch = fakeFetch({ message: "not found" }, 404);
    await assert.rejects(fetchRate("USD", "XYZ", "", fetch), /No rate is available/);
  });

  test("error when the server fails", async () => {
    const fetch = fakeFetch({}, 500);
    await assert.rejects(fetchRate("USD", "INR", "", fetch), /error \(500\)/);
  });

  test("error when there is no internet", async () => {
    const offline = async () => {
      throw new TypeError("Failed to fetch");
    };
    await assert.rejects(fetchRate("USD", "INR", "", offline), /internet connection/);
  });
});

describe("fetchRateHistory", () => {
  test("returns points sorted oldest first", async () => {
    const fetch = fakeFetch({
      amount: 1,
      base: "USD",
      start_date: "2024-03-11",
      end_date: "2024-03-13",
      rates: { "2024-03-13": { INR: 82.9 }, "2024-03-11": { INR: 82.7 }, "2024-03-12": { INR: 82.8 } },
    });
    const points = await fetchRateHistory("USD", "INR", "2024-03-11", "2024-03-13", fetch);
    assert.equal(fetch.lastUrl, "https://api.frankfurter.dev/v1/2024-03-11..2024-03-13?base=USD&symbols=INR");
    assert.deepEqual(points, [
      { date: "2024-03-11", rate: 82.7 },
      { date: "2024-03-12", rate: 82.8 },
      { date: "2024-03-13", rate: 82.9 },
    ]);
  });
});

describe("fetchCurrencies", () => {
  test("returns code and name, sorted", async () => {
    const fetch = fakeFetch({ USD: "United States Dollar", EUR: "Euro", INR: "Indian Rupee" });
    const list = await fetchCurrencies(fetch);
    assert.deepEqual(
      list.map((c) => c.code),
      ["EUR", "INR", "USD"]
    );
    assert.equal(list[1].name, "Indian Rupee");
  });
});

describe("convert (everything together)", () => {
  test("100 USD → INR at 84.12 = 8412.00", async () => {
    const fetch = fakeFetch({ amount: 1, base: "USD", date: "2026-10-21", rates: { INR: 84.12 } });
    const result = await convert({ amount: "100", from: "USD", to: "INR", date: "" }, fetch);
    assert.deepEqual(result, {
      amount: 100,
      from: "USD",
      to: "INR",
      rate: 84.12,
      rateDate: "2026-10-21",
      result: 8412,
    });
  });

  test("bad input is rejected before calling the API", async () => {
    const fetch = fakeFetch({});
    await assert.rejects(convert({ amount: "-1", from: "USD", to: "USD" }, fetch), ConversionError);
    assert.equal(fetch.lastUrl, undefined);
  });
});
