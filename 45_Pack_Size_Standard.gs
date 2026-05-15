/////////////////////////////////////
// PACK SIZE STANDARD
/////////////////////////////////////

function parsePackSizeStandard_(input) {
	const source = normalisePackSizeInput_(input);

	const result = {
		rawPackSize: source.rawPackSize,
		rawCaseSize: source.rawCaseSize,
		cleanedPackSize: "",
		displayPackSize: "",
		packQty: "",
		baseUnit: "",
		unitPerPackCase: "",
		reviewFlag: "OK",
		notes: "",
	};

	let raw = cleanPackSizeStandardText_(source.rawPackSize);
	let caseSize = cleanPackSizeStandardText_(source.rawCaseSize);

	if (!raw && caseSize) {
		raw = caseSize;
		caseSize = "";
	}

	if (!raw) {
		result.reviewFlag = "CHECK PACK SIZE";
		result.notes = "Empty pack size";
		return result;
	}

	if (caseSize) {
		raw = caseSize + "x" + raw;
	}

	raw = applyPackSizeOcrFixes_(raw);

	result.cleanedPackSize = raw;

	const parsed = parseCleanPackSizeStandard_(raw);

	if (!parsed.ok) {
		result.reviewFlag = "CHECK PACK SIZE";
		result.notes =
			parsed.notes || "Unrecognised pack size format: " + result.rawPackSize;
		return result;
	}

	result.displayPackSize = parsed.displayPackSize;
	result.packQty = parsed.packQty;
	result.baseUnit = parsed.baseUnit;
	result.unitPerPackCase = parsed.unitPerPackCase;
	result.notes = parsed.notes || "";

	return result;
}

/////////////////////////////////////
// NORMALISE INPUT
/////////////////////////////////////

function normalisePackSizeInput_(input) {
	if (input && typeof input === "object") {
		return {
			rawPackSize: input.packSize || input.pack_size || "",
			rawCaseSize: input.caseSize || input.case_size || "",
		};
	}

	return {
		rawPackSize: input || "",
		rawCaseSize: "",
	};
}

/////////////////////////////////////
// CLEAN PACK SIZE TEXT
/////////////////////////////////////

function cleanPackSizeStandardText_(value) {
	return String(value || "")
		.toLowerCase()
		.replace(/\s+/g, "")
		.replace(/\u00d7/g, "x")
		.replace(/\u2013/g, "-")
		.replace(/\u2014/g, "-")
		.replace(/litres/g, "ltr")
		.replace(/litre/g, "ltr")
		.replace(/liters/g, "ltr")
		.replace(/liter/g, "ltr")
		.replace(/ltrs\b/g, "ltr")
		.replace(/lt\b/g, "ltr")
		.replace(/(\d+(?:\.\d+)?)l\b/g, "$1ltr")
		.replace(/dozens/g, "dozen")
		.replace(/bags/g, "bag")
		.trim();
}

/////////////////////////////////////
// OCR FIXES
/////////////////////////////////////

function applyPackSizeOcrFixes_(raw) {
	let text = (raw || "").toString().toLowerCase().trim();

	const exactFixes = {
		"1-51tr": "1-5ltr",
		"2-51tr": "2-5ltr",
		"2-2.271tr": "2-2.27ltr",
		"1-6rol1": "1-6roll",
		"12-50oml": "12-500ml",
		"24-50oml": "24-500ml",
		"1-2opk": "1-20pk",
		"200-99": "200-9g",
		"300-49": "300-4g",
		"100-29": "100-2g",

		// Bidfood OCR / formatting fixes
		"10-1.00pk": "10-100pk",
		"1-1.80pk": "1-180pk",
	};

	if (exactFixes[text]) return exactFixes[text];

	text = text.replace(/\s+/g, "");
	text = text.replace(/[.,]+$/g, "");

	if (exactFixes[text]) return exactFixes[text];

	text = text.replace(/rol1/g, "roll");
	text = text.replace(/m1\b/g, "ml");
	text = text.replace(/(\d+)o(ml|g|kg|ltr|pk|ea|ptn)$/g, "$10$2");

	text = text.replace(/(\d+(?:\.\d+)?)1tr\b/g, "$1ltr");
	text = text.replace(/^(\d+)1t\b/g, "$1ltr");
	text = text.replace(/\b11t\b/g, "1ltr");

	text = text.replace(/(\d+)\.\.(\d+)(kg|g|ml|ltr|l)\b/g, "$1.$2$3");

	text = text.replace(/(\d+)x(\d+)p\b/g, "$1x$2g");
	text = text.replace(/-(\d+)p\b/g, "-$1g");

	text = text.replace(/(\d+)l\b/g, "$1ltr");

	return text;
}

/////////////////////////////////////
// PARSE CLEAN PACK SIZE
/////////////////////////////////////

function parseCleanPackSizeStandard_(raw) {
	let match;

	/////////////////////////////////////
	// dozen
	/////////////////////////////////////

	if (raw === "dozen") {
		return {
			ok: true,
			displayPackSize: "dozen",
			packQty: 12,
			baseUnit: "each",
			unitPerPackCase: 12,
			notes: "",
		};
	}

	/////////////////////////////////////
	// 15dozen
	/////////////////////////////////////

	match = raw.match(/^(\d+(?:\.\d+)?)dozen$/);

	if (match) {
		const qty = Number(match[1]) * 12;

		return {
			ok: true,
			displayPackSize: match[1] + "dozen",
			packQty: qty,
			baseUnit: "each",
			unitPerPackCase: qty,
			notes: "",
		};
	}

	/////////////////////////////////////
	// 2x15dozen
	/////////////////////////////////////

	match = raw.match(/^(\d+)x(\d+(?:\.\d+)?)dozen$/);

	if (match) {
		const outer = Number(match[1]);
		const dozen = Number(match[2]);
		const qty = outer * dozen * 12;

		return {
			ok: true,
			displayPackSize: outer + "x" + dozen + "dozen",
			packQty: outer,
			baseUnit: "each",
			unitPerPackCase: qty,
			notes: "",
		};
	}

	/////////////////////////////////////
	// 24x2x28.5g / 5-90x20g
	/////////////////////////////////////

	match = raw.match(/^(\d+)[x-](\d+)x(\d+(?:\.\d+)?)(g|kg|ml|ltr|m)$/);

	if (match) {
		const outer = Number(match[1]);
		const inner = Number(match[2]);
		const unitSize = Number(match[3]);
		const converted = convertPackUnitTotal_(inner * unitSize, match[4]);

		return {
			ok: true,
			displayPackSize:
				outer + "x" + inner + "x" + unitSize + displayUnit_(match[4]),
			packQty: outer,
			baseUnit: converted.baseUnit,
			unitPerPackCase: outer * converted.total,
			notes: "",
		};
	}

	/////////////////////////////////////
	// 2-2.27ltr / 2-5l / 24x330ml / 6-2.62kg
	/////////////////////////////////////

	match = raw.match(/^(\d+)[x-](\d+(?:\.\d+)?)(g|kg|ml|ltr|m)$/);

	if (match) {
		const packQty = Number(match[1]);
		const unitSize = Number(match[2]);
		const converted = convertPackUnitTotal_(packQty * unitSize, match[3]);

		return {
			ok: true,
			displayPackSize: packQty + "x" + unitSize + displayUnit_(match[3]),
			packQty: packQty,
			baseUnit: converted.baseUnit,
			unitPerPackCase: converted.total,
			notes: "",
		};
	}

	/////////////////////////////////////
	// 25-170-200
	// Portion weight range, e.g. 25 portions at 170-200g each
	/////////////////////////////////////

	match = raw.match(/^(\d+)[x-](\d+)-(\d+)$/);

	if (match) {
		const qty = Number(match[1]);
		const minWeight = Number(match[2]);
		const maxWeight = Number(match[3]);

		return {
			ok: true,
			displayPackSize: qty + "x" + minWeight + "-" + maxWeight + "g",
			packQty: qty,
			baseUnit: "each",
			unitPerPackCase: qty,
			notes:
				"Portion weight range retained: " + minWeight + "-" + maxWeight + "g",
		};
	}

	/////////////////////////////////////
	// 1-120pk / 1-500ea / 6-100ptn / 1-6roll
	/////////////////////////////////////

	match = raw.match(
		/^(\d+)[x-](\d+)(pk|ea|each|unit|units|ptn|ptns|portion|portions|roll|rolls|sti|stick|sticks|can|cans|btl|btls|sac|sachet|sachets|box|boxes)$/,
	);

	if (match) {
		const outer = Number(match[1]);
		const inner = Number(match[2]);
		const qty = outer * inner;
		const packWord = normaliseEachPackWord_(match[3]);

		return {
			ok: true,
			displayPackSize:
				outer === 1 ? inner + packWord : outer + "x" + inner + packWord,
			packQty: qty,
			baseUnit: "each",
			unitPerPackCase: qty,
			notes: "",
		};
	}

	/////////////////////////////////////
	// 120pk / 500ea / 2000sac
	/////////////////////////////////////

	match = raw.match(
		/^(\d+)(pk|ea|each|unit|units|ptn|ptns|portion|portions|roll|rolls|sti|stick|sticks|can|cans|btl|btls|sac|sachet|sachets|box|boxes)$/,
	);

	if (match) {
		const qty = Number(match[1]);

		return {
			ok: true,
			displayPackSize: qty + normaliseEachPackWord_(match[2]),
			packQty: qty,
			baseUnit: "each",
			unitPerPackCase: qty,
			notes: "",
		};
	}

	/////////////////////////////////////
	// 48x9inch
	/////////////////////////////////////

	match = raw.match(/^(\d+)x(\d+(?:\.\d+)?)(inch|in)$/);

	if (match) {
		const qty = Number(match[1]);

		return {
			ok: true,
			displayPackSize: qty + "x" + match[2] + "inch",
			packQty: qty,
			baseUnit: "each",
			unitPerPackCase: qty,
			notes: "Size descriptor retained: " + match[2] + "inch",
		};
	}

	/////////////////////////////////////
	// 5kg / 1.8kg / 500ml / 5l
	/////////////////////////////////////

	match = raw.match(/^(\d+(?:\.\d+)?)(g|kg|ml|ltr|m)$/);

	if (match) {
		const unitSize = Number(match[1]);
		const converted = convertPackUnitTotal_(unitSize, match[2]);

		return {
			ok: true,
			displayPackSize: unitSize + displayUnit_(match[2]),
			packQty: 1,
			baseUnit: converted.baseUnit,
			unitPerPackCase: converted.total,
			notes: "",
		};
	}

	/////////////////////////////////////
	// 2000
	/////////////////////////////////////

	match = raw.match(/^(\d+)$/);

	if (match) {
		const qty = Number(match[1]);

		return {
			ok: true,
			displayPackSize: String(qty),
			packQty: qty,
			baseUnit: "each",
			unitPerPackCase: qty,
			notes: "",
		};
	}

	return {
		ok: false,
		notes: "Unrecognised pack size format: " + raw,
	};
}

/////////////////////////////////////
// UNIT CONVERSION
/////////////////////////////////////

function convertPackUnitTotal_(total, unit) {
	if (unit === "kg") {
		return {
			baseUnit: "g",
			total: total * 1000,
		};
	}

	if (unit === "ltr") {
		return {
			baseUnit: "ml",
			total: total * 1000,
		};
	}

	return {
		baseUnit: unit,
		total: total,
	};
}

/////////////////////////////////////
// DISPLAY UNIT
/////////////////////////////////////

function displayUnit_(unit) {
	if (unit === "ltr") return "ltr";
	return unit;
}

/////////////////////////////////////
// NORMALISE EACH WORDS
/////////////////////////////////////

function normaliseEachPackWord_(word) {
	const value = String(word || "").toLowerCase();

	if (value === "each") return "ea";
	if (value === "unit" || value === "units") return "ea";
	if (value === "ptns" || value === "portion" || value === "portions")
		return "ptn";
	if (value === "rolls") return "roll";
	if (value === "sticks") return "sti";
	if (value === "stick") return "sti";
	if (value === "cans") return "can";
	if (value === "btls") return "btl";
	if (value === "sachet" || value === "sachets") return "sac";
	if (value === "boxes") return "box";

	return value;
}

/////////////////////////////////////
// COMPATIBILITY WRAPPER
// MATCHES OLD parsePackSizeToUnits_ OUTPUT SHAPE
/////////////////////////////////////

function parsePackSizeToUnitsStandard_(packSize) {
	const parsed = parsePackSizeStandard_(packSize);

	return {
		packQty: parsed.packQty,
		baseUnit: parsed.baseUnit,
		unitPerCase: parsed.unitPerPackCase,
		unitPerPackCase: parsed.unitPerPackCase,
		reviewFlag: parsed.reviewFlag,
		notes: parsed.notes,
		displayPackSize: parsed.displayPackSize,
		cleanedPackSize: parsed.cleanedPackSize,
	};
}

/////////////////////////////////////
// PILGRIM PACK SIZE STANDARD WRAPPER
/////////////////////////////////////

function buildPilgrimStandardPackSize_(caseSize, packSize) {
	const parsed = parsePackSizeStandard_({
		caseSize: caseSize,
		packSize: packSize,
	});

	return {
		pack_size: parsed.displayPackSize,
		packQty: parsed.packQty,
		baseUnit: parsed.baseUnit,
		unitPerPackCase: parsed.unitPerPackCase,
		reviewFlag: parsed.reviewFlag,
		notes: parsed.notes,
	};
}

/////////////////////////////////////
// BIDFOOD PACK SIZE STANDARD WRAPPER
/////////////////////////////////////

function buildBidfoodStandardPackSize_(packSize) {
	const parsed = parsePackSizeStandard_(packSize);

	return {
		pack_size: parsed.displayPackSize,
		packQty: parsed.packQty,
		baseUnit: parsed.baseUnit,
		unitPerPackCase: parsed.unitPerPackCase,
		reviewFlag: parsed.reviewFlag,
		notes: parsed.notes,
	};
}
