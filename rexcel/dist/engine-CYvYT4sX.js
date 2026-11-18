import { NS as e, Package as t, esc as n, escText as r, formats as i, newPackage as a } from "@voffice/react-shared";
// src/engine/model.ts
var o = (e) => {
	let t = "";
	for (e++; e > 0;) {
		let n = (e - 1) % 26;
		t = String.fromCharCode(65 + n) + t, e = Math.floor((e - 1) / 26);
	}
	return t;
}, s = (e) => {
	let t = 0;
	for (let n of e.toUpperCase()) t = t * 26 + (n.charCodeAt(0) - 64);
	return t - 1;
}, c = (e, t) => o(t) + (e + 1);
function l(e) {
	let t = /^([A-Za-z]+)(\d+)$/.exec(e);
	return t ? {
		r: t[2] - 1,
		c: s(t[1])
	} : null;
}
var u = (e) => ({
	name: e,
	cells: /* @__PURE__ */ new Map(),
	merges: [],
	colWidths: /* @__PURE__ */ new Map(),
	rowHeights: /* @__PURE__ */ new Map(),
	defaultColWidth: 64,
	defaultRowHeight: 17
}), d = () => ({
	sheets: [u("Sheet1")],
	styles: [{}],
	active: 0
});
function f(e, t) {
	let n = JSON.stringify(t), r = e.styles.findIndex((e) => JSON.stringify(e) === n);
	return r >= 0 ? r : (e.styles.push(t), e.styles.length - 1);
}
function p(e, t, n) {
	return e.cells.get(c(t, n));
}
function m(e, t, n, r) {
	let i = c(t, n);
	!r || r.v == null && !r.f && r.s == null ? e.cells.delete(i) : e.cells.set(i, r);
}
function h(e, t, n) {
	if (!e) return "";
	if (e.t === "e") return String(e.v);
	let r = e.s == null ? void 0 : t.styles[e.s];
	return n(e.v, r?.numFmt);
}
var g = (e) => ({
	sheets: e.sheets.map((e) => ({
		...e,
		cells: new Map(e.cells),
		merges: e.merges.map((e) => ({ ...e })),
		colWidths: new Map(e.colWidths),
		rowHeights: new Map(e.rowHeights)
	})),
	styles: e.styles.map((e) => ({ ...e })),
	active: e.active
});
// src/engine/xlsx-read.ts
e.s;
var ee = e.r, _ = (e, t) => e.children.filter((e) => e.localName === t), v = (e, t) => e.children.find((e) => e.localName === t), y = (e, t) => e ? e.getAttribute(t) : null, te = {
	0: "General",
	1: "0",
	2: "0.00",
	3: "#,##0",
	4: "#,##0.00",
	5: "#,##0;[Red]-#,##0",
	9: "0%",
	10: "0.00%",
	11: "0.00E+00",
	12: "# ?/?",
	13: "# ??/??",
	14: "m/d/yy",
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm AM/PM",
	19: "h:mm:ss AM/PM",
	20: "h:mm",
	21: "h:mm:ss",
	22: "m/d/yy h:mm",
	37: "#,##0;(#,##0)",
	38: "#,##0;[Red](#,##0)",
	39: "#,##0.00;(#,##0.00)",
	40: "#,##0.00;[Red](#,##0.00)",
	44: "_(\"$\"* #,##0.00_)",
	45: "mm:ss",
	46: "[h]:mm:ss",
	47: "mmss.0",
	48: "##0.0E+0",
	49: "@"
};
function b(e) {
	let t = v(e, "t");
	return t ? t.textContent : _(e, "r").map((e) => v(e, "t")?.textContent || "").join("");
}
function ne(e) {
	let t = {
		fonts: [],
		fills: [],
		borders: [],
		xfs: [],
		numFmts: /* @__PURE__ */ new Map()
	};
	if (!e.has("xl/styles.xml")) return t;
	let n = e.xml("xl/styles.xml"), r = v(n, "numFmts");
	if (r) for (let e of _(r, "numFmt")) t.numFmts.set(+(y(e, "numFmtId") || 0), y(e, "formatCode") || "General");
	let i = v(n, "fonts");
	i && (t.fonts = _(i, "font"));
	let a = v(n, "fills");
	a && (t.fills = _(a, "fill"));
	let o = v(n, "borders");
	o && (t.borders = _(o, "border"));
	let s = v(n, "cellXfs");
	return s && (t.xfs = _(s, "xf")), t;
}
function x(e) {
	if (!e) return;
	let t = y(e, "rgb");
	if (t && t.length === 8) return ("#" + t.slice(2)).toLowerCase();
	if (t && t.length === 6) return ("#" + t).toLowerCase();
}
function S(e, t) {
	let n = {};
	if (!e) return n;
	let r = t.fonts[+(y(e, "fontId") || 0)];
	if (r) {
		let e = {}, t = v(r, "sz");
		t && (e.size = +(y(t, "val") || 11));
		let i = v(r, "name");
		i && (e.name = y(i, "val") || void 0), v(r, "b") && (e.bold = !0), v(r, "i") && (e.italic = !0), v(r, "u") && (e.underline = !0);
		let a = x(v(r, "color"));
		a && (e.color = a), Object.keys(e).length && (n.font = e);
	}
	let i = t.fills[+(y(e, "fillId") || 0)];
	if (i) {
		let e = v(i, "patternFill"), t = x(v(e || i, "fgColor"));
		t && y(e, "patternType") !== "none" && (n.fill = t);
	}
	let a = t.borders[+(y(e, "borderId") || 0)];
	if (a) {
		let e = {};
		for (let t of [
			"left",
			"right",
			"top",
			"bottom"
		]) {
			let n = v(a, t);
			n && y(n, "style") && (e[t] = x(v(n, "color")) || "#000000");
		}
		Object.keys(e).length && (n.border = e);
	}
	let o = v(e, "alignment");
	if (o) {
		let e = {}, t = y(o, "horizontal");
		t && (e.h = t === "general" ? void 0 : t);
		let r = y(o, "vertical");
		r && (e.v = r === "center" ? "middle" : r), y(o, "wrapText") === "1" && (e.wrap = !0), Object.keys(e).length && (n.align = e);
	}
	let s = +(y(e, "numFmtId") || 0), c = t.numFmts.get(s) || te[s];
	return c && c !== "General" && (n.numFmt = c), n;
}
async function C(e) {
	let n = await t.open(e), r = n.xml("xl/workbook.xml"), i = n.relsOf("xl/workbook.xml"), a = [];
	if (n.has("xl/sharedStrings.xml")) {
		let e = n.xml("xl/sharedStrings.xml");
		for (let t of _(e, "si")) a.push(b(t));
	}
	let o = ne(n), s = {
		sheets: [],
		styles: [{}],
		active: 0
	}, c = [], l = v(r, "sheets"), d = l ? _(l, "sheet") : [];
	for (let e of d) {
		let t = y(e, "name") || "Sheet", r = e.getAttributeNS(ee, "id") || y(e, "r:id") || "", l = i.find((e) => e.id === r), d = l ? l.target.startsWith("/") ? l.target.slice(1) : "xl/" + l.target.replace(/^\.\//, "") : "";
		c.push(d);
		let f = u(t);
		d && n.has(d) && w(n, d, f, a, o, s), s.sheets.push(f);
	}
	s.sheets.length || (s.sheets.push(u("Sheet1")), c.push("xl/worksheets/sheet1.xml"));
	let f = y(v(r, "bookViews") && v(v(r, "bookViews"), "workbookView"), "activeTab");
	return f && (s.active = Math.min(+f, s.sheets.length - 1)), {
		wb: s,
		pkg: n,
		sheetParts: c
	};
}
function w(e, t, n, r, i, a) {
	let o = e.xml(t), s = v(o, "cols");
	if (s) for (let e of _(s, "col")) {
		let t = +(y(e, "min") || 1), r = +(y(e, "max") || t), i = +(y(e, "width") || 8.43);
		for (let e = t; e <= r && e <= 16384; e++) n.colWidths.set(e - 1, Math.round(i * 7 + 5));
	}
	let u = v(o, "sheetData");
	if (u) for (let e of _(u, "row")) {
		let t = (y(e, "r") || 1) - 1, o = y(e, "ht");
		o && n.rowHeights.set(t, Math.round(+o));
		for (let t of _(e, "c")) {
			let e = l(y(t, "r") || "");
			if (!e) continue;
			let o = y(t, "t") || "n", s = y(t, "s"), u = v(t, "v"), d = v(t, "f"), p = {};
			if (d && (p.f = d.textContent), o === "s") p.v = r[+(u?.textContent || 0)] ?? "", p.t = "s";
			else if (o === "inlineStr") {
				let e = v(t, "is");
				p.v = e ? b(e) : "", p.t = "s";
			} else if (o === "str") p.v = u?.textContent || "", p.t = "s";
			else if (o === "b") p.v = u?.textContent === "1", p.t = "b";
			else if (o === "e") p.v = u?.textContent || "#VALUE!", p.t = "e";
			else {
				let e = u?.textContent;
				p.v = e != null && e !== "" ? +e : null, p.t = "n";
			}
			if (s != null) {
				let e = S(i.xfs[+s], i);
				Object.keys(e).length && (p.s = f(a, e));
			}
			(p.v != null || p.f || p.s != null) && n.cells.set(c(e.r, e.c), p);
		}
	}
	let d = v(o, "mergeCells");
	if (d) for (let e of _(d, "mergeCell")) {
		let [t, r] = (y(e, "ref") || "").split(":"), i = l(t), a = l(r || t);
		i && a && n.merges.push({
			r1: i.r,
			c1: i.c,
			r2: a.r,
			c2: a.c
		});
	}
	let p = v(o, "sheetViews"), m = p && v(v(p, "sheetView") || p, "pane");
	if (m) {
		let e = +(y(m, "xSplit") || 0), t = +(y(m, "ySplit") || 0);
		(e || t) && (n.frozen = {
			rows: t,
			cols: e
		});
	}
}
// src/engine/xlsx-write.ts
var T = `xmlns="${e.s}" xmlns:r="${e.r}"`;
function E(e) {
	return Math.max(0, Math.round((e - 5) / 7 * 100) / 100);
}
function re(e, t) {
	let n = e.strings.get(t);
	return n ?? (n = e.strList.length, e.strList.push(t), e.strings.set(t, n)), n;
}
function ie(e, t, n, i) {
	let a = c(t, n), o = i.s ? ` s="${i.s}"` : "";
	return i.f == null ? i.t === "e" ? `<c r="${a}"${o} t="e"><v>${r(String(i.v))}</v></c>` : i.t === "b" ? `<c r="${a}"${o} t="b"><v>${+!!i.v}</v></c>` : typeof i.v == "string" ? `<c r="${a}"${o} t="s"><v>${re(e, i.v)}</v></c>` : typeof i.v == "number" ? `<c r="${a}"${o}><v>${i.v}</v></c>` : "" : `<c r="${a}"${o}><f>${r(i.f)}</f></c>`;
}
function D(e, t) {
	let n = /* @__PURE__ */ new Map(), r = 0, i = 0;
	for (let [e, a] of t.cells) {
		let t = l(e);
		r = Math.max(r, t.r), i = Math.max(i, t.c);
		let o = n.get(t.r) || [];
		o.push({
			r: t.r,
			c: t.c,
			cell: a
		}), n.set(t.r, o);
	}
	let a = `A1:${o(i)}${r + 1}`, s = "";
	t.colWidths.size && (s = "<cols>" + [...t.colWidths.entries()].sort((e, t) => e[0] - t[0]).map(([e, t]) => `<col min="${e + 1}" max="${e + 1}" width="${E(t)}" customWidth="1"/>`).join("") + "</cols>");
	let u = [];
	for (let [r, i] of [...n.entries()].sort((e, t) => e[0] - t[0])) {
		i.sort((e, t) => e.c - t.c);
		let n = t.rowHeights.get(r), a = n == null ? "" : ` ht="${n}" customHeight="1"`;
		u.push(`<row r="${r + 1}"${a}>` + i.map((t) => ie(e, t.r, t.c, t.cell)).join("") + "</row>");
	}
	let d = "";
	t.merges.length && (d = `<mergeCells count="${t.merges.length}">` + t.merges.map((e) => `<mergeCell ref="${c(e.r1, e.c1)}:${c(e.r2, e.c2)}"/>`).join("") + "</mergeCells>");
	let f = "";
	if (t.frozen && (t.frozen.rows || t.frozen.cols)) {
		let { rows: e, cols: n } = t.frozen;
		f = `<pane xSplit="${n}" ySplit="${e}" topLeftCell="${c(e, n)}" activePane="bottomRight" state="frozen"/>`;
	}
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet ${T}><dimension ref="${a}"/><sheetViews><sheetView workbookViewId="0">${f}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>${s}<sheetData>${u.join("")}</sheetData>${d}</worksheet>`;
}
function O(e) {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<sst ${T} count="${e.strList.length}" uniqueCount="${e.strList.length}">` + e.strList.map((e) => `<si><t xml:space="preserve">${r(e)}</t></si>`).join("") + "</sst>";
}
function ae(e) {
	let t = [], r = ["<fill><patternFill patternType=\"none\"/></fill>", "<fill><patternFill patternType=\"gray125\"/></fill>"], i = ["<border><left/><right/><top/><bottom/><diagonal/></border>"], a = [], o = /* @__PURE__ */ new Map(), s = [], c = /* @__PURE__ */ new Map(), l = /* @__PURE__ */ new Map(), u = /* @__PURE__ */ new Map(), d = 164;
	for (let f of e.styles) {
		let e = 0;
		if (f.font) {
			let r = JSON.stringify(f.font);
			if (!c.has(r)) {
				let e = f.font, i = "";
				e.bold && (i += "<b/>"), e.italic && (i += "<i/>"), e.underline && (i += "<u/>"), i += `<sz val="${e.size || 11}"/>`, e.color && (i += `<color rgb="FF${e.color.replace("#", "").toUpperCase()}"/>`), i += `<name val="${n(e.name || "Calibri")}"/>`, c.set(r, t.length), t.push(`<font>${i}</font>`);
			}
			e = c.get(r);
		}
		let p = 0;
		if (f.fill) {
			let e = f.fill;
			l.has(e) || (l.set(e, r.length), r.push(`<fill><patternFill patternType="solid"><fgColor rgb="FF${e.replace("#", "").toUpperCase()}"/><bgColor indexed="64"/></patternFill></fill>`)), p = l.get(e);
		}
		let m = 0;
		if (f.border) {
			let e = JSON.stringify(f.border);
			if (!u.has(e)) {
				let t = f.border, n = "";
				for (let e of [
					"left",
					"right",
					"top",
					"bottom"
				]) n += t[e] ? `<${e} style="thin"><color rgb="FF${t[e].replace("#", "").toUpperCase()}"/></${e}>` : `<${e}/>`;
				u.set(e, i.length), i.push(`<border>${n}<diagonal/></border>`);
			}
			m = u.get(e);
		}
		let h = 0;
		f.numFmt && f.numFmt !== "General" && (o.has(f.numFmt) || (o.set(f.numFmt, d), a.push(`<numFmt numFmtId="${d}" formatCode="${n(f.numFmt)}"/>`), d++), h = o.get(f.numFmt));
		let g = "";
		if (f.align) {
			let e = f.align;
			g = `<alignment${e.h ? ` horizontal="${e.h}"` : ""}${e.v ? ` vertical="${e.v === "middle" ? "center" : e.v}"` : ""}${e.wrap ? " wrapText=\"1\"" : ""}/>`;
		}
		s.push(`<xf numFmtId="${h}" fontId="${e}" fillId="${p}" borderId="${m}" xfId="0"${h ? " applyNumberFormat=\"1\"" : ""}${e ? " applyFont=\"1\"" : ""}${p ? " applyFill=\"1\"" : ""}${m ? " applyBorder=\"1\"" : ""}${g ? " applyAlignment=\"1\"" : ""}>${g}</xf>`);
	}
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<styleSheet ${T}>` + (a.length ? `<numFmts count="${a.length}">${a.join("")}</numFmts>` : "") + `<fonts count="${t.length || 1}">${t.join("") || "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>"}</fonts><fills count="${r.length}">${r.join("")}</fills><borders count="${i.length}">${i.join("")}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${s.length}">${s.join("")}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
}
var oe = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", se = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", ce = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
async function le(t, r = {}) {
	let o = i[r.format || "xlsx"] || i.xlsx, s = {
		wb: t,
		strings: /* @__PURE__ */ new Map(),
		strList: []
	}, c = t.sheets.map((e) => D(s, e)), l = r.pkg || a("xl/workbook.xml", "", o), u = r.sheetParts && r.sheetParts.length === t.sheets.length ? r.sheetParts : t.sheets.map((e, t) => `xl/worksheets/sheet${t + 1}.xml`);
	c.forEach((e, t) => l.addPart(u[t], e, ce)), l.addPart("xl/sharedStrings.xml", O(s), oe), l.addPart("xl/styles.xml", ae(t), se);
	let d = r.pkg ? ue(r.pkg) : null, f = t.sheets.map((e) => e.name);
	if (d && d.length === f.length && d.every((e, t) => e === f[t])) r.pkg;
	else {
		let r = t.sheets.map((e, t) => `<sheet name="${n(e.name)}" sheetId="${t + 1}" r:id="rId${t + 1}"/>`).join("");
		l.setPart("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<workbook ${T}><bookViews><workbookView activeTab="${t.active}"/></bookViews><sheets>${r}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`);
		let i = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"];
		t.sheets.forEach((t, n) => i.push(`<Relationship Id="rId${n + 1}" Type="${e.rel}/worksheet" Target="worksheets/sheet${n + 1}.xml"/>`));
		let a = t.sheets.length;
		i.push(`<Relationship Id="rId${a + 1}" Type="${e.rel}/styles" Target="styles.xml"/>`), i.push(`<Relationship Id="rId${a + 2}" Type="${e.rel}/sharedStrings" Target="sharedStrings.xml"/>`), i.push("</Relationships>"), l.setPart("xl/_rels/workbook.xml.rels", i.join(""));
	}
	return r.pkg || l.addRel("", e.rel + "/officeDocument", "xl/workbook.xml"), l.save(o.mime);
}
function ue(e) {
	if (!e.has("xl/workbook.xml")) return null;
	let t = e.xml("xl/workbook.xml").children.find((e) => e.localName === "sheets");
	return t ? t.children.filter((e) => e.localName === "sheet").map((e) => e.getAttribute("name") || "") : null;
}
// src/engine/csv.ts
function de(e, t = ",") {
	let n = d(), r = n.sheets[0];
	return r.name = "Sheet1", fe(e, t).forEach((e, t) => {
		e.forEach((e, n) => {
			if (e === "") return;
			let i = Number(e), a = e.trim() !== "" && !isNaN(i) && /^[-+]?[\d.]+(e[-+]?\d+)?$/i.test(e.trim());
			r.cells.set(c(t, n), a ? {
				v: i,
				t: "n"
			} : {
				v: e,
				t: "s"
			});
		});
	}), n;
}
function fe(e, t) {
	let n = [], r = [], i = "", a = !1;
	for (let o = 0; o < e.length; o++) {
		let s = e[o];
		a ? s === "\"" ? e[o + 1] === "\"" ? (i += "\"", o++) : a = !1 : i += s : s === "\"" ? a = !0 : s === t ? (r.push(i), i = "") : s === "\n" ? (r.push(i), n.push(r), r = [], i = "") : s === "\r" || (i += s);
	}
	return (i !== "" || r.length) && (r.push(i), n.push(r)), n;
}
function pe(e, t, n = ",") {
	let r = 0, i = 0;
	for (let e of t.cells.keys()) {
		let t = /^([A-Z]+)(\d+)$/.exec(e), n = 0;
		for (let e of t[1]) n = n * 26 + (e.charCodeAt(0) - 64);
		i = Math.max(i, n - 1), r = Math.max(r, t[2] - 1);
	}
	let a = [];
	for (let e = 0; e <= r; e++) {
		let r = [];
		for (let a = 0; a <= i; a++) {
			let i = t.cells.get(c(e, a)), o = i ? String(i.v ?? "") : "";
			(o.includes(n) || o.includes("\"") || o.includes("\n")) && (o = "\"" + o.replace(/"/g, "\"\"") + "\""), r.push(o);
		}
		a.push(r.join(n));
	}
	return a.join("\r\n");
}
// src/engine/formula.ts
var k = class {
	code;
	constructor(e) {
		this.code = e;
	}
	toString() {
		return this.code;
	}
}, A = () => new k("#DIV/0!"), j = () => new k("#VALUE!"), M = () => new k("#REF!"), me = () => new k("#NAME?"), he = () => new k("#N/A"), N = () => new k("#NUM!");
function ge(e) {
	let t = [], n = 0, r = e.length;
	for (; n < r;) {
		let i = e[n];
		if (i === " " || i === "	" || i === "\n") n++;
		else if (i === "\"") {
			let i = "";
			for (n++; n < r;) if (e[n] === "\"") {
				if (e[n + 1] === "\"") i += "\"", n += 2;
				else {
					n++;
					break;
				}
			} else i += e[n++];
			t.push({
				t: "str",
				v: i
			});
		} else if (/[0-9.]/.test(i)) {
			let i = "";
			for (; n < r && /[0-9.]/.test(e[n]);) i += e[n++];
			if (e[n] === "e" || e[n] === "E") for (i += e[n++], (e[n] === "+" || e[n] === "-") && (i += e[n++]); n < r && /[0-9]/.test(e[n]);) i += e[n++];
			t.push({
				t: "num",
				v: i
			});
		} else if (/[A-Za-z_$]/.test(i)) {
			let i = "";
			for (; n < r && /[A-Za-z0-9_.$!]/.test(e[n]);) i += e[n++];
			t.push({
				t: "ident",
				v: i
			});
		} else if (i === "<" && e[n + 1] === ">") t.push({
			t: "op",
			v: "<>"
		}), n += 2;
		else if (i === "<" && e[n + 1] === "=") t.push({
			t: "op",
			v: "<="
		}), n += 2;
		else if (i === ">" && e[n + 1] === "=") t.push({
			t: "op",
			v: ">="
		}), n += 2;
		else if ("+-*/^&=<>%,():".includes(i)) t.push({
			t: i === "(" ? "lp" : i === ")" ? "rp" : i === "," ? "comma" : i === ":" ? "colon" : "op",
			v: i
		}), n++;
		else throw j();
	}
	return t;
}
function P(e) {
	let t = e.toUpperCase();
	if (t === "TRUE") return {
		k: "bool",
		v: !0
	};
	if (t === "FALSE") return {
		k: "bool",
		v: !1
	};
	if (/^#(DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!|NULL!)$/.test(t)) return {
		k: "err",
		v: t
	};
	let n, r = e, i = e.lastIndexOf("!");
	i >= 0 && (n = e.slice(0, i).replace(/^'(.*)'$/, "$1"), r = e.slice(i + 1));
	let a = r.replace(/\$/g, ""), o = /^([A-Za-z]+)(\d+)$/.exec(a);
	return o ? {
		k: "ref",
		ref: {
			r: o[2] - 1,
			c: s(o[1]),
			sheet: n
		}
	} : null;
}
var F = {
	"=": 1,
	"<>": 1,
	"<": 1,
	">": 1,
	"<=": 1,
	">=": 1,
	"&": 2,
	"+": 3,
	"-": 3,
	"*": 4,
	"/": 4,
	"^": 5
};
function _e(e) {
	let t = ge(e), n = 0, r = () => t[n], i = () => t[n++];
	function a(e) {
		let t = o();
		for (;;) {
			let n = r();
			if (!n || n.t !== "op" || !(n.v in F) || F[n.v] < e) break;
			let o = i().v, s = a(F[o] + (o === "^" ? 0 : 1));
			t = {
				k: "bin",
				op: o,
				a: t,
				b: s
			};
		}
		return t;
	}
	function o() {
		let e = r();
		return e && e.t === "op" && (e.v === "-" || e.v === "+") ? (i(), {
			k: "un",
			op: e.v,
			a: o()
		}) : s();
	}
	function s() {
		let e = c();
		for (; r() && r().t === "op" && r().v === "%";) i(), e = {
			k: "pct",
			a: e
		};
		return e;
	}
	function c() {
		let e = i();
		if (!e) throw j();
		if (e.t === "num") return {
			k: "num",
			v: parseFloat(e.v)
		};
		if (e.t === "str") return {
			k: "str",
			v: e.v
		};
		if (e.t === "lp") {
			let e = a(1);
			if (i()?.t !== "rp") throw j();
			return e;
		}
		if (e.t === "ident") {
			if (r() && r().t === "lp") {
				i();
				let t = [];
				if (r() && r().t !== "rp") for (; t.push(a(1)), r()?.t === "comma";) i();
				if (i()?.t !== "rp") throw j();
				return {
					k: "call",
					name: e.v.toUpperCase(),
					args: t
				};
			}
			if (r() && r().t === "colon") {
				i();
				let t = i(), n = P(e.v), r = t && t.t === "ident" ? P(t.v) : null;
				if (n && n.k === "ref" && r && r.k === "ref") return {
					k: "range",
					range: {
						r1: Math.min(n.ref.r, r.ref.r),
						c1: Math.min(n.ref.c, r.ref.c),
						r2: Math.max(n.ref.r, r.ref.r),
						c2: Math.max(n.ref.c, r.ref.c),
						sheet: n.ref.sheet
					}
				};
				throw M();
			}
			let t = P(e.v);
			if (t) return t;
			throw me();
		}
		throw j();
	}
	let l = a(1);
	if (n < t.length) throw j();
	return l;
}
function I(e) {
	if (e instanceof k) return e;
	if (Array.isArray(e)) return e.length ? I(e[0]) : j();
	if (typeof e == "number") return e;
	if (typeof e == "boolean") return +!!e;
	if (e === "") return 0;
	let t = Number(e);
	return isNaN(t) ? j() : t;
}
function L(e) {
	return e instanceof k ? e : Array.isArray(e) ? e.length ? L(e[0]) : j() : typeof e == "boolean" ? e ? "TRUE" : "FALSE" : String(e);
}
function R(e) {
	if (e instanceof k) return e;
	if (Array.isArray(e)) return e.length ? R(e[0]) : j();
	if (typeof e == "boolean") return e;
	if (typeof e == "number") return e !== 0;
	if (e === "TRUE") return !0;
	if (e === "FALSE") return !1;
	let t = Number(e);
	return isNaN(t) ? j() : t !== 0;
}
var z = (e) => e instanceof k;
function B(e) {
	return Array.isArray(e) ? e.flatMap(B) : [e];
}
function V(e) {
	let t = [];
	for (let n of B(e)) typeof n == "number" ? t.push(n) : typeof n == "string" && n !== "" && !isNaN(Number(n)) ? t.push(Number(n)) : typeof n == "boolean" && t.push(+!!n);
	return t;
}
function H(e, t) {
	switch (e.k) {
		case "num": return e.v;
		case "str": return e.v;
		case "bool": return e.v;
		case "err": return new k(e.v);
		case "ref": return t.get(e.ref);
		case "range": return t.range(e.range);
		case "pct": {
			let n = I(H(e.a, t));
			return z(n) ? n : n / 100;
		}
		case "un": {
			let n = I(H(e.a, t));
			return z(n) ? n : e.op === "-" ? -n : n;
		}
		case "bin": {
			let n = H(e.a, t), r = H(e.b, t);
			if (z(n)) return n;
			if (z(r)) return r;
			if (e.op === "&") {
				let e = L(n), t = L(r);
				return z(e) ? e : z(t) ? t : e + t;
			}
			if ([
				"=",
				"<>",
				"<",
				">",
				"<=",
				">="
			].includes(e.op)) return U(n, r, e.op);
			let i = I(n), a = I(r);
			if (z(i)) return i;
			if (z(a)) return a;
			switch (e.op) {
				case "+": return i + a;
				case "-": return i - a;
				case "*": return i * a;
				case "/": return a === 0 ? A() : i / a;
				case "^": {
					let e = i ** +a;
					return isNaN(e) ? N() : e;
				}
			}
			return j();
		}
		case "call": return Te(e.name, e.args.map((e) => H(e, t)), t);
	}
}
function U(e, t, n) {
	let r = typeof e == "number", i = typeof t == "number", a = typeof e == "string", o = typeof t == "string", s = typeof e == "boolean", c = typeof t == "boolean", l;
	if (r && i) l = e - t;
	else if (a && o) l = e.localeCompare(t);
	else if (s && c) l = e === t ? 0 : e ? 1 : -1;
	else {
		let n = (e) => typeof e == "number" ? 0 : typeof e == "string" ? 1 : 2;
		l = n(e) - n(t);
	}
	switch (n) {
		case "=": return l === 0;
		case "<>": return l !== 0;
		case "<": return l < 0;
		case ">": return l > 0;
		case "<=": return l <= 0;
		case ">=": return l >= 0;
	}
	return j();
}
var W = (e) => {
	for (let t of B(e)) if (z(t)) return t;
	return null;
}, ve = {
	SUM: (e) => W(e) || V(e).reduce((e, t) => e + t, 0),
	AVERAGE: (e) => {
		let t = W(e);
		if (t) return t;
		let n = V(e);
		return n.length ? n.reduce((e, t) => e + t, 0) / n.length : A();
	},
	MIN: (e) => {
		let t = V(e);
		return t.length ? Math.min(...t) : 0;
	},
	MAX: (e) => {
		let t = V(e);
		return t.length ? Math.max(...t) : 0;
	},
	COUNT: (e) => V(e).length,
	COUNTA: (e) => B(e).filter((e) => e !== "" && e != null).length,
	COUNTBLANK: (e) => B(e).filter((e) => e === "" || e == null).length,
	PRODUCT: (e) => {
		let t = V(e);
		return t.length ? t.reduce((e, t) => e * t, 1) : 0;
	},
	SUMPRODUCT: (e) => {
		let t = e.map((e) => B(e)), n = Math.max(...t.map((e) => e.length), 0), r = 0;
		for (let e = 0; e < n; e++) {
			let n = 1;
			for (let r of t) {
				let t = I(r[e] ?? 0);
				if (z(t)) return t;
				n *= t;
			}
			r += n;
		}
		return r;
	},
	ROUND: (e) => {
		let t = I(e[0]), n = I(e[1] ?? 0);
		if (z(t) || z(n)) return j();
		let r = 10 ** n;
		return Math.round(t * r) / r;
	},
	ROUNDUP: (e) => {
		let t = I(e[0]), n = I(e[1] ?? 0);
		if (z(t) || z(n)) return j();
		let r = 10 ** n;
		return (t < 0 ? -1 : 1) * Math.ceil(Math.abs(t) * r) / r;
	},
	ROUNDDOWN: (e) => {
		let t = I(e[0]), n = I(e[1] ?? 0);
		if (z(t) || z(n)) return j();
		let r = 10 ** n;
		return (t < 0 ? -1 : 1) * Math.floor(Math.abs(t) * r) / r;
	},
	INT: (e) => {
		let t = I(e[0]);
		return z(t) ? t : Math.floor(t);
	},
	MOD: (e) => {
		let t = I(e[0]), n = I(e[1]);
		return z(t) || z(n) ? j() : n === 0 ? A() : t - n * Math.floor(t / n);
	},
	ABS: (e) => {
		let t = I(e[0]);
		return z(t) ? t : Math.abs(t);
	},
	SIGN: (e) => {
		let t = I(e[0]);
		return z(t) ? t : Math.sign(t);
	},
	POWER: (e) => {
		let t = I(e[0]), n = I(e[1]);
		if (z(t) || z(n)) return j();
		let r = t ** +n;
		return isNaN(r) ? N() : r;
	},
	SQRT: (e) => {
		let t = I(e[0]);
		return z(t) ? t : t < 0 ? N() : Math.sqrt(t);
	},
	EXP: (e) => {
		let t = I(e[0]);
		return z(t) ? t : Math.exp(t);
	},
	LN: (e) => {
		let t = I(e[0]);
		return z(t) ? t : t <= 0 ? N() : Math.log(t);
	},
	LOG: (e) => {
		let t = I(e[0]), n = I(e[1] ?? 10);
		return z(t) || z(n) ? j() : t <= 0 ? N() : Math.log(t) / Math.log(n);
	},
	LOG10: (e) => {
		let t = I(e[0]);
		return z(t) ? t : t <= 0 ? N() : Math.log10(t);
	},
	SIN: (e) => {
		let t = I(e[0]);
		return z(t) ? t : Math.sin(t);
	},
	COS: (e) => {
		let t = I(e[0]);
		return z(t) ? t : Math.cos(t);
	},
	TAN: (e) => {
		let t = I(e[0]);
		return z(t) ? t : Math.tan(t);
	},
	PI: () => Math.PI,
	RAND: () => Math.random(),
	RANDBETWEEN: (e) => {
		let t = I(e[0]), n = I(e[1]);
		return z(t) || z(n) ? j() : Math.floor(t + Math.random() * (n - t + 1));
	},
	TODAY: () => q(/* @__PURE__ */ new Date()),
	NOW: () => q(/* @__PURE__ */ new Date()) + we(/* @__PURE__ */ new Date()),
	DATE: (e) => {
		let t = I(e[0]), n = I(e[1]), r = I(e[2]);
		return z(t) || z(n) || z(r) ? j() : q(new Date(t, n - 1, r));
	},
	YEAR: (e) => {
		let t = I(e[0]);
		return z(t) ? t : J(t).getFullYear();
	},
	MONTH: (e) => {
		let t = I(e[0]);
		return z(t) ? t : J(t).getMonth() + 1;
	},
	DAY: (e) => {
		let t = I(e[0]);
		return z(t) ? t : J(t).getDate();
	},
	HOUR: (e) => {
		let t = I(e[0]);
		return z(t) ? t : Math.floor(t % 1 * 24);
	},
	MINUTE: (e) => {
		let t = I(e[0]);
		return z(t) ? t : Math.floor(t % 1 * 24 % 1 * 60);
	},
	SECOND: (e) => {
		let t = I(e[0]);
		return z(t) ? t : Math.round(t % 1 * 24 % 1 * 60 % 1 * 60);
	},
	WEEKDAY: (e) => {
		let t = I(e[0]);
		return z(t) ? t : J(t).getDay() + 1;
	},
	LEFT: (e) => {
		let t = L(e[0]), n = I(e[1] ?? 1);
		return z(t) || z(n) ? j() : t.slice(0, n);
	},
	RIGHT: (e) => {
		let t = L(e[0]), n = I(e[1] ?? 1);
		return z(t) || z(n) ? j() : n === 0 ? "" : t.slice(-n);
	},
	MID: (e) => {
		let t = L(e[0]), n = I(e[1]), r = I(e[2]);
		return z(t) || z(n) || z(r) ? j() : t.substr(n - 1, r);
	},
	LEN: (e) => {
		let t = L(e[0]);
		return z(t) ? t : t.length;
	},
	TRIM: (e) => {
		let t = L(e[0]);
		return z(t) ? t : t.replace(/\s+/g, " ").trim();
	},
	UPPER: (e) => {
		let t = L(e[0]);
		return z(t) ? t : t.toUpperCase();
	},
	LOWER: (e) => {
		let t = L(e[0]);
		return z(t) ? t : t.toLowerCase();
	},
	PROPER: (e) => {
		let t = L(e[0]);
		return z(t) ? t : t.replace(/\w\S*/g, (e) => e[0].toUpperCase() + e.slice(1).toLowerCase());
	},
	SUBSTITUTE: (e) => {
		let t = L(e[0]), n = L(e[1]), r = L(e[2]);
		return z(t) || z(n) || z(r) ? j() : t.split(n).join(r);
	},
	REPLACE: (e) => {
		let t = L(e[0]), n = I(e[1]), r = I(e[2]), i = L(e[3]);
		return z(t) || z(n) || z(r) || z(i) ? j() : t.slice(0, n - 1) + i + t.slice(n - 1 + r);
	},
	FIND: (e) => {
		let t = L(e[0]), n = L(e[1]), r = I(e[2] ?? 1);
		if (z(t) || z(n) || z(r)) return j();
		let i = n.indexOf(t, r - 1);
		return i < 0 ? j() : i + 1;
	},
	SEARCH: (e) => {
		let t = L(e[0]), n = L(e[1]), r = I(e[2] ?? 1);
		if (z(t) || z(n) || z(r)) return j();
		let i = n.toLowerCase().indexOf(t.toLowerCase(), r - 1);
		return i < 0 ? j() : i + 1;
	},
	VALUE: (e) => {
		let t = L(e[0]);
		if (z(t)) return t;
		let n = Number(t);
		return isNaN(n) ? j() : n;
	},
	CONCATENATE: (e) => W(e) || B(e).map((e) => L(e)).filter((e) => !z(e)).join(""),
	CONCAT: (e) => ve.CONCATENATE(e, null),
	EXACT: (e) => {
		let t = L(e[0]), n = L(e[1]);
		return z(t) || z(n) ? j() : t === n;
	},
	REPT: (e) => {
		let t = L(e[0]), n = I(e[1]);
		return z(t) || z(n) ? j() : t.repeat(Math.max(0, Math.floor(n)));
	},
	AND: (e) => W(e) || B(e).every((e) => R(e) === !0),
	OR: (e) => W(e) || B(e).some((e) => R(e) === !0),
	NOT: (e) => {
		let t = R(e[0]);
		return z(t) ? t : !t;
	},
	XOR: (e) => W(e) || B(e).filter((e) => R(e) === !0).length % 2 == 1,
	IF: (e) => {
		let t = R(e[0]);
		return z(t) ? t : t ? e[1] ?? !0 : e[2] ?? !1;
	},
	IFERROR: (e) => z(e[0]) ? e[1] ?? "" : e[0],
	ISBLANK: (e) => e[0] === "" || e[0] == null,
	ISNUMBER: (e) => typeof e[0] == "number",
	ISTEXT: (e) => typeof e[0] == "string",
	ISERROR: (e) => z(e[0]),
	ISNA: (e) => z(e[0]) && e[0].code === "#N/A",
	ISLOGICAL: (e) => typeof e[0] == "boolean",
	CHOOSE: (e) => {
		let t = I(e[0]);
		return z(t) ? t : e[t] ?? j();
	},
	VLOOKUP: (e) => Ce(e, G, !0),
	HLOOKUP: (e) => Ce(e, G, !1),
	INDEX: (e) => {
		let t = B(e[0]), n = I(e[1] ?? 1), r = I(e[2] ?? 1);
		if (z(n) || z(r)) return j();
		let i = e[0].__dims, a = i ? i[1] : 1;
		return t[(n - 1) * a + (r - 1)] ?? M();
	},
	MATCH: (e) => {
		let t = e[0], n = B(e[1]);
		I(e[2] ?? 1);
		for (let e = 0; e < n.length; e++) if (U(n[e], t, "=") === !0) return e + 1;
		return he();
	},
	SUMIF: (e) => be(e, !1),
	COUNTIF: (e) => ye(e),
	AVERAGEIF: (e) => be(e, !0),
	SUMIFS: (e) => xe(e),
	COUNTIFS: (e) => Se(e),
	ROW: (e) => e[0] == null && G.current ? G.current.r + 1 : 0,
	COLUMN: (e) => e[0] == null && G.current ? G.current.c + 1 : 0,
	ROWS: (e) => {
		let t = e[0].__dims;
		return t ? t[0] : Array.isArray(e[0]) ? e[0].length : 1;
	},
	COLUMNS: (e) => {
		let t = e[0].__dims;
		return t ? t[1] : 1;
	},
	OFFSET: (e) => M(),
	INDIRECT: (e) => M(),
	TRANSPOSE: (e) => B(e[0]),
	TEXT: (e) => {
		let t = e[0], n = L(e[1]);
		return z(n) ? n : String(t);
	}
}, G = {
	get: () => M(),
	range: () => []
};
function K(e, t) {
	if (typeof t == "string") {
		let n = /^(<=|>=|<>|<|>|=)(.*)$/.exec(t);
		return n ? U(e, isNaN(Number(n[2])) ? n[2] : Number(n[2]), n[1] === "=" ? "=" : n[1]) === !0 : t.includes("*") || t.includes("?") ? RegExp("^" + t.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i").test(String(e)) : U(e, isNaN(Number(t)) ? t : Number(t), "=") === !0;
	}
	return U(e, t, "=") === !0;
}
function ye(e) {
	let t = B(e[0]), n = e[1], r = 0;
	for (let e of t) K(e, n) && r++;
	return r;
}
function be(e, t) {
	let n = B(e[0]), r = e[1], i = e[2] ? B(e[2]) : n, a = 0, o = 0;
	for (let e = 0; e < n.length; e++) if (K(n[e], r)) {
		let t = I(i[e] ?? 0);
		z(t) || (a += t, o++);
	}
	return t ? o ? a / o : A() : a;
}
function xe(e) {
	let t = B(e[0]), n = 0, r = [];
	for (let t = 1; t + 1 < e.length + 1; t += 2) r.push([B(e[t]), e[t + 1]]);
	for (let e = 0; e < t.length; e++) if (r.every(([t, n]) => K(t[e], n))) {
		let r = I(t[e]);
		z(r) || (n += r);
	}
	return n;
}
function Se(e) {
	let t = [];
	for (let n = 0; n + 1 < e.length + 1; n += 2) t.push([B(e[n]), e[n + 1]]);
	let n = t[0]?.[0].length || 0, r = 0;
	for (let e = 0; e < n; e++) t.every(([t, n]) => K(t[e], n)) && r++;
	return r;
}
function Ce(e, t, n) {
	let r = e[0], i = e[1], a = I(e[2]), o = e[3] == null || R(e[3]);
	if (z(a)) return j();
	let s = B(i), c = i.__dims, l = c ? c[0] : s.length, u = c ? c[1] : 1;
	if (n) {
		let e = -1;
		for (let t = 0; t < l; t++) {
			let n = s[t * u];
			if (U(n, r, "=") === !0) {
				e = t;
				break;
			}
			o === !0 && U(n, r, "<") === !0 && (e = t);
		}
		return e < 0 ? he() : s[e * u + (a - 1)] ?? M();
	}
	{
		let e = -1;
		for (let t = 0; t < u; t++) {
			let n = s[t];
			if (U(n, r, "=") === !0) {
				e = t;
				break;
			}
			o === !0 && U(n, r, "<") === !0 && (e = t);
		}
		return e < 0 ? he() : s[(a - 1) * u + e] ?? M();
	}
}
function q(e) {
	let t = Date.UTC(1899, 11, 30), n = Date.UTC(e.getFullYear(), e.getMonth(), e.getDate()), r = Math.round((n - t) / 864e5);
	return r >= 60 && r++, r;
}
function J(e) {
	let t = Math.floor(e);
	return t >= 60 && t--, new Date(Date.UTC(1899, 11, 30) + t * 864e5);
}
function we(e) {
	return (e.getHours() * 3600 + e.getMinutes() * 60 + e.getSeconds()) / 86400;
}
function Te(e, t, n) {
	G = n;
	let r = ve[e];
	if (!r) return me();
	try {
		return r(t, n);
	} catch (e) {
		return e instanceof k ? e : j();
	}
}
function Y(e, t) {
	switch (e.k) {
		case "ref":
			t.refs.push(e.ref);
			break;
		case "range":
			t.ranges.push(e.range);
			break;
		case "un":
		case "pct":
			Y(e.a, t);
			break;
		case "bin":
			Y(e.a, t), Y(e.b, t);
			break;
		case "call": e.args.forEach((e) => Y(e, t));
	}
}
// src/engine/calc.ts
function X(e, t, n = 0) {
	if (!t) return n;
	let r = e.sheets.findIndex((e) => e.name === t);
	return r >= 0 ? r : n;
}
function Ee(e) {
	let t = [], n = /* @__PURE__ */ new Map();
	e.sheets.forEach((r, i) => {
		for (let [a, o] of r.cells) {
			if (!o.f) continue;
			let r = l(a);
			try {
				let s = _e(o.f), l = {
					refs: [],
					ranges: []
				};
				Y(s, l);
				let u = [];
				for (let t of l.refs) u.push(X(e, t.sheet, i) + "!" + c(t.r, t.c));
				for (let t of l.ranges) {
					let n = X(e, t.sheet, i);
					for (let e = t.r1; e <= t.r2; e++) for (let r = t.c1; r <= t.c2; r++) u.push(n + "!" + c(e, r));
				}
				let d = {
					key: i + "!" + a,
					sheet: i,
					r: r.r,
					c: r.c,
					ast: s,
					deps: u
				};
				t.push(d), n.set(d.key, d);
			} catch {
				o.v = "#VALUE!", o.t = "e";
			}
		}
	});
	let r = /* @__PURE__ */ new Map(), i = /* @__PURE__ */ new Map();
	for (let e of t) {
		let t = 0;
		for (let r of e.deps) n.has(r) && (t++, i.set(r, [...i.get(r) || [], e.key]));
		r.set(e.key, t);
	}
	let a = t.filter((e) => (r.get(e.key) || 0) === 0).map((e) => e.key), o = [];
	for (; a.length;) {
		let e = a.shift();
		o.push(e);
		for (let t of i.get(e) || []) {
			let e = (r.get(t) || 0) - 1;
			r.set(t, e), e === 0 && a.push(t);
		}
	}
	let s = new Set(t.filter((e) => !o.includes(e.key)).map((e) => e.key)), u = (t, n, r) => {
		let i = e.sheets[t]?.cells.get(c(n, r));
		return i ? i.v ?? "" : "";
	};
	for (let t of o) {
		let r = n.get(t), i = e.sheets[r.sheet].cells.get(c(r.r, r.c)), a = {
			current: {
				r: r.r,
				c: r.c
			},
			get: (t) => u(X(e, t.sheet, r.sheet), t.r, t.c),
			range: (t) => {
				let n = X(e, t.sheet, r.sheet), i = [];
				for (let e = t.r1; e <= t.r2; e++) for (let r = t.c1; r <= t.c2; r++) i.push(u(n, e, r));
				return i.__dims = [t.r2 - t.r1 + 1, t.c2 - t.c1 + 1], i;
			}
		}, o;
		try {
			o = H(r.ast, a);
		} catch (e) {
			o = e instanceof k ? e : new k("#VALUE!");
		}
		Array.isArray(o) && (o = o.length ? o[0] : ""), o instanceof k ? (i.v = o.code, i.t = "e") : typeof o == "number" ? (i.v = o, i.t = "n") : typeof o == "boolean" ? (i.v = o, i.t = "b") : (i.v = typeof o == "string" ? o : String(o), i.t = "s");
	}
	for (let t of s) {
		let r = n.get(t), i = e.sheets[r.sheet].cells.get(c(r.r, r.c));
		i.v = "#REF!", i.t = "e";
	}
}
// src/engine/numfmt.ts
var De = [
	"Jan",
	"Feb",
	"Mar",
	"Apr",
	"May",
	"Jun",
	"Jul",
	"Aug",
	"Sep",
	"Oct",
	"Nov",
	"Dec"
], Oe = [
	"January",
	"February",
	"March",
	"April",
	"May",
	"June",
	"July",
	"August",
	"September",
	"October",
	"November",
	"December"
], ke = [
	"Sun",
	"Mon",
	"Tue",
	"Wed",
	"Thu",
	"Fri",
	"Sat"
];
function Ae(e, t) {
	if (e == null) return "";
	if (typeof e == "boolean") return e ? "TRUE" : "FALSE";
	if (typeof e == "string") return e;
	let n = t || "General";
	return n === "General" ? je(e) : n === "@" ? String(e) : Me(n) ? Ne(e, n) : Pe(e, n);
}
function je(e) {
	if (!isFinite(e) || Number.isInteger(e) && Math.abs(e) < 1e11) return String(e);
	let t = e.toPrecision(11);
	return String(parseFloat(t));
}
function Me(e) {
	return !/[#0]/.test(e) && /[ymdhs]/i.test(e);
}
function Z(e, t) {
	return String(e).padStart(t, "0");
}
function Ne(e, t) {
	let n = J(e), r = e - Math.floor(e), i = Math.round(r * 86400), a = Math.floor(i / 3600), o = Math.floor(i % 3600 / 60), s = i % 60, c = n.getUTCFullYear(), l = n.getUTCMonth(), u = n.getUTCDate(), d = n.getUTCDay(), f = "", p = 0, m = t;
	for (; p < m.length;) {
		let e = m[p], n = (e) => {
			let t = 0;
			for (; m[p + t] === e;) t++;
			return t;
		};
		if (e === "y" || e === "Y") {
			let t = n(e);
			p += t, f += t >= 3 ? String(c) : Z(c % 100, 2);
		} else if (e === "m" || e === "M") {
			let t = n(e);
			p += t, f.slice(-1), /h$/i.test(m.slice(0, p - t)) || /^s/i.test(m.slice(p)) ? f += t >= 2 ? Z(o, 2) : String(o) : f += t >= 4 ? Oe[l] : t === 3 ? De[l] : t === 2 ? Z(l + 1, 2) : String(l + 1);
		} else if (e === "d" || e === "D") {
			let t = n(e);
			p += t, f += t >= 4 || t === 3 ? ke[d] : t === 2 ? Z(u, 2) : String(u);
		} else if (e === "h" || e === "H") {
			let r = n(e);
			p += r;
			let i = /AM\/PM/i.test(t) ? a % 12 || 12 : a;
			f += r >= 2 ? Z(i, 2) : String(i);
		} else if (e === "s" || e === "S") {
			let t = n(e);
			p += t, f += t >= 2 ? Z(s, 2) : String(s);
		} else e === "A" && /^AM\/PM/i.test(m.slice(p)) ? (f += a < 12 ? "AM" : "PM", p += 5) : e === "a" && /^am\/pm/i.test(m.slice(p)) ? (f += a < 12 ? "am" : "pm", p += 5) : (f += e, p++);
	}
	return f;
}
function Pe(e, t) {
	let n = t, r = !1;
	n.includes("%") && (r = !0, e *= 100, n = n.replace(/%/g, "")), n = n.replace(/\[[^\]]*\]/g, ""), n = n.split(";")[0];
	let i = n.indexOf("."), a = i >= 0 ? n.slice(i + 1) : "", o = (a.match(/[0#]/g) || []).length, s = /,/.test(i >= 0 ? n.slice(0, i) : n), c = ((i >= 0 ? n.slice(0, i) : n).match(/^[^0#,]*/) || [""])[0].replace(/"/g, ""), l = (a.match(/[^0#]*$/) || [""])[0].replace(/"/g, "") + (r ? "%" : ""), u = Math.abs(e).toFixed(o);
	if (s) {
		let [e, t] = u.split(".");
		u = e.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (t ? "." + t : "");
	}
	return (e < 0 ? "-" : "") + c + u + l;
}
// src/engine/grid.ts
var Fe = 46, Ie = 22;
function Le(e, t = 200, n = 2e3) {
	let r = [46];
	for (let n = 0; n < t; n++) r.push(r[n] + (e.colWidths.get(n) ?? e.defaultColWidth));
	let i = [22];
	for (let t = 0; t < n; t++) i.push(i[t] + (e.rowHeights.get(t) ?? e.defaultRowHeight));
	return {
		xs: r,
		ys: i,
		nw: t,
		nh: n
	};
}
function Re(e, t, n, r) {
	let i = t.merges.find((e) => n >= e.r1 && n <= e.r2 && r >= e.c1 && r <= e.c2);
	return i ? {
		x: e.xs[i.c1],
		y: e.ys[i.r1],
		w: e.xs[i.c2 + 1] - e.xs[i.c1],
		h: e.ys[i.r2 + 1] - e.ys[i.r1]
	} : {
		x: e.xs[r],
		y: e.ys[n],
		w: e.xs[r + 1] - e.xs[r],
		h: e.ys[n + 1] - e.ys[n]
	};
}
function ze(e, t, n) {
	if (t < 46 || n < 22) return null;
	let r = 0;
	for (; r < e.nw && e.xs[r + 1] < t;) r++;
	let i = 0;
	for (; i < e.nh && e.ys[i + 1] < n;) i++;
	return {
		r: i,
		c: r
	};
}
var Q = (e, t) => getComputedStyle(document.documentElement).getPropertyValue(e).trim() || t;
function Be(e, t) {
	let n = window.devicePixelRatio || 1, r = e.clientWidth, i = e.clientHeight;
	(e.width !== r * n || e.height !== i * n) && (e.width = r * n, e.height = i * n);
	let a = e.getContext("2d");
	a.setTransform(n * t.zoom, 0, 0, n * t.zoom, 0, 0), a.clearRect(0, 0, r / t.zoom, i / t.zoom);
	let { layout: s, sheet: l, wb: u, sel: d } = t, f = 46 - t.scrollX, p = 22 - t.scrollY, m = r / t.zoom, h = i / t.zoom, g = Q("--sheet-grid", "#c0c0c0"), ee = Q("--hdr-a", "#fff"), _ = Q("--hdr-b", "#E4ECF7"), v = Q("--hdr-line", "#9EB6CE"), y = Q("--hdr-ink", "#000"), te = Q("--cell-sel", "rgba(49,74,160,0.2)"), b = Q("--cell-sel-border", "#000"), ne = Math.max(0, $(s.xs, t.scrollX + 46) - 1), x = Math.min(s.nw - 1, $(s.xs, t.scrollX + m) + 1), S = Math.max(0, $(s.ys, t.scrollY + 22) - 1), C = Math.min(s.nh - 1, $(s.ys, t.scrollY + h) + 1);
	a.save(), a.beginPath(), a.rect(46, 22, m - 46, h - 22), a.clip(), a.translate(f, p), a.textBaseline = "middle";
	for (let e = S; e <= C; e++) for (let n = ne; n <= x; n++) {
		let r = Re(s, l, e, n), i = l.cells.get(c(e, n)), o = i?.s == null ? void 0 : u.styles[i.s];
		if (o?.fill && (a.fillStyle = o.fill, a.fillRect(r.x, r.y, r.w, r.h)), t.showGrid && (a.strokeStyle = g, a.lineWidth = 1 / t.zoom, a.strokeRect(r.x + .5, r.y + .5, r.w, r.h)), o?.border) {
			a.strokeStyle = "#000", a.lineWidth = 1 / t.zoom;
			let e = o.border;
			a.beginPath(), e.top && (a.moveTo(r.x, r.y + .5), a.lineTo(r.x + r.w, r.y + .5)), e.bottom && (a.moveTo(r.x, r.y + r.h - .5), a.lineTo(r.x + r.w, r.y + r.h - .5)), e.left && (a.moveTo(r.x + .5, r.y), a.lineTo(r.x + .5, r.y + r.h)), e.right && (a.moveTo(r.x + r.w - .5, r.y), a.lineTo(r.x + r.w - .5, r.y + r.h)), a.stroke();
		}
		if (i && i.v != null && i.v !== "") {
			let e = o?.font, t = e?.size || 11;
			a.font = `${e?.italic ? "italic " : ""}${e?.bold ? "bold " : ""}${t}px ${e?.name || "Calibri"}, Carlito, sans-serif`, a.fillStyle = e?.color || Q("--ink", "#000");
			let n = i.t === "e" ? String(i.v) : Ae(i.v, o?.numFmt), s = o?.align?.h || (typeof i.v == "number" ? "right" : "left"), c = r.x + 3;
			a.textAlign = "left", s === "right" ? (c = r.x + r.w - 3, a.textAlign = "right") : s === "center" && (c = r.x + r.w / 2, a.textAlign = "center"), a.save(), a.beginPath(), a.rect(r.x, r.y, r.w, r.h), a.clip(), a.fillText(n, c, r.y + r.h / 2), a.restore();
		}
	}
	let w = Ve(d), T = s.xs[w.c1], E = s.ys[w.r1], re = s.xs[w.c2 + 1] - T, ie = s.ys[w.r2 + 1] - E;
	a.fillStyle = te, a.fillRect(T, E, re, ie), a.strokeStyle = b, a.lineWidth = 2 / t.zoom, a.strokeRect(T + 1, E + 1, re - 2, ie - 2);
	let D = Re(s, l, d.r, d.c);
	a.strokeStyle = b, a.lineWidth = 2 / t.zoom, a.strokeRect(D.x + 1, D.y + 1, D.w - 2, D.h - 2), a.restore(), a.save();
	let O = a.createLinearGradient(0, 0, 0, 22);
	O.addColorStop(0, ee), O.addColorStop(1, _), a.fillStyle = O, a.fillRect(0, 0, m, 22), a.fillRect(0, 0, 46, h), a.strokeStyle = v, a.lineWidth = 1 / t.zoom, a.font = `11px ${Q("--ui-font", "Tahoma")}`, a.fillStyle = y, a.textAlign = "center", a.textBaseline = "middle";
	for (let e = ne; e <= x; e++) {
		let t = s.xs[e] + f, n = s.xs[e + 1] - s.xs[e];
		e >= w.c1 && e <= w.c2 && (a.fillStyle = Q("--hdr-sel-b", "#F8C55E"), a.fillRect(t, 0, n, 22), a.fillStyle = y), a.strokeRect(t + .5, .5, n, 22), a.fillText(o(e), t + n / 2, 11);
	}
	a.textAlign = "center";
	for (let e = S; e <= C; e++) {
		let t = s.ys[e] + p, n = s.ys[e + 1] - s.ys[e];
		e >= w.r1 && e <= w.r2 && (a.fillStyle = Q("--hdr-sel-b", "#F8C55E"), a.fillRect(0, t, 46, n), a.fillStyle = y), a.strokeRect(.5, t + .5, 46, n), a.fillText(String(e + 1), 23, t + n / 2);
	}
	a.fillStyle = O, a.fillRect(0, 0, 46, 22), a.strokeRect(.5, .5, 46, 22), a.restore();
}
function $(e, t) {
	let n = 0, r = e.length - 1;
	for (; n < r;) {
		let i = n + r >> 1;
		e[i] < t ? n = i + 1 : r = i;
	}
	return n;
}
function Ve(e) {
	return {
		r1: Math.min(e.r, e.r2),
		c1: Math.min(e.c, e.c2),
		r2: Math.max(e.r, e.r2),
		c2: Math.max(e.c, e.c2)
	};
}
export { m as C, l as S, s as _, Be as a, u as b, Ae as c, pe as d, le as f, g, h, Le as i, Ee as l, c as m, Fe as n, Ve as o, C as p, Re as r, ze as s, Ie as t, de as u, o as v, f as w, d as x, p as y };
