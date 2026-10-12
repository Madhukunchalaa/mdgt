"use client";

/**
 * CMT Code editor: one 2-digit numeric code for every value of a characteristic.
 *
 *   <CmtCodeEditor values={["FLANGE","PIPE"]} codes={{FLANGE:"01"}} onChange={setCodes} />
 *
 * Rules (same as the backend, matg_attributes/catalog.py):
 *   - mandatory for every value
 *   - exactly 2 digits, 0-9 only
 *   - not repeated inside the same characteristic
 */

export const CMT_CODE_REGEX = /^[0-9]{2}$/;

/** Returns a list of readable problems ([] when everything is fine). */
export function validateCmtCodes(values = [], codes = {}) {
  const problems = [];
  const seen = {};
  for (const value of values) {
    const code = (codes[value] ?? "").toString();
    if (!code) {
      problems.push(`CMT Code is required for "${value}"`);
    } else if (!CMT_CODE_REGEX.test(code)) {
      problems.push(`CMT Code for "${value}" must be exactly 2 digits`);
    } else if (seen[code] !== undefined) {
      problems.push(`CMT Code ${code} is used for both "${seen[code]}" and "${value}"`);
    } else {
      seen[code] = value;
    }
  }
  return problems;
}

/** Keep only codes for values that still exist. */
export function pickCodes(values = [], codes = {}) {
  const out = {};
  for (const value of values) {
    if (codes[value] !== undefined && codes[value] !== "") out[value] = codes[value];
  }
  return out;
}

export default function CmtCodeEditor({ values = [], codes = {}, onChange, disabled = false }) {
  if (!values || values.length === 0) return null;

  const problems = validateCmtCodes(values, codes);
  const codeCount = {};
  values.forEach((v) => {
    const c = codes[v];
    if (c) codeCount[c] = (codeCount[c] || 0) + 1;
  });

  const setCode = (value, raw) => {
    // digits only, max 2 characters
    const clean = raw.replace(/[^0-9]/g, "").slice(0, 2);
    onChange({ ...codes, [value]: clean });
  };

  return (
    <div className="mt-2 border rounded-lg bg-gray-50 p-2">
      <div className="flex items-center justify-between mb-1">
        <label className="block text-xs font-medium text-gray-700">
          CMT Code * <span className="font-normal text-gray-500">(2 digits, numbers only)</span>
        </label>
        {problems.length === 0 ? (
          <span className="text-xs text-green-600">All codes set</span>
        ) : (
          <span className="text-xs text-amber-600">{problems.length} to fix</span>
        )}
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-1.5 max-h-48 overflow-y-auto">
        {values.map((value) => {
          const code = codes[value] ?? "";
          const invalid = code !== "" && !CMT_CODE_REGEX.test(code);
          const repeated = code !== "" && codeCount[code] > 1;
          const missing = code === "";
          const bad = invalid || repeated;
          return (
            <div key={value} className="flex items-center gap-2">
              <span className="flex-1 truncate text-xs text-gray-800" title={value}>
                {value}
              </span>
              <input
                type="text"
                inputMode="numeric"
                maxLength={2}
                disabled={disabled}
                value={code}
                onChange={(e) => setCode(value, e.target.value)}
                placeholder="00"
                aria-label={`CMT Code for ${value}`}
                className={`w-14 px-2 py-1 text-sm text-center font-mono border rounded-md focus:outline-none focus:ring-2 ${
                  bad
                    ? "border-red-500 focus:ring-red-500"
                    : missing
                    ? "border-amber-400 focus:ring-amber-400"
                    : "focus:ring-blue-500"
                }`}
              />
            </div>
          );
        })}
      </div>
      {problems.length > 0 && (
        <p className="text-xs text-red-600 mt-1">{problems[0]}</p>
      )}
    </div>
  );
}
