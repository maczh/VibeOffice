import { NS as e, Package as t, esc as n, escText as r, formats as i, newPackage as a } from "@voffice/react-shared";
// src/engine/model.ts
var o = 0, s = (e = "b") => `${e}${++o}_${Math.random().toString(36).slice(2, 7)}`, c = (e) => ({
	kind: "p",
	id: s("p"),
	runs: [],
	pPr: e
}), l = () => ({
	pageW: 12240,
	pageH: 15840,
	margins: {
		top: 1440,
		right: 1440,
		bottom: 1440,
		left: 1440
	}
}), u = () => ({
	blocks: [c()],
	section: l(),
	styles: [
		{
			id: "Normal",
			name: "Normal"
		},
		{
			id: "Heading1",
			name: "Heading 1",
			rPr: {
				bold: !0,
				size: 16
			},
			pPr: {
				spaceBefore: 12,
				spaceAfter: 6
			}
		},
		{
			id: "Heading2",
			name: "Heading 2",
			rPr: {
				bold: !0,
				size: 13
			},
			pPr: {
				spaceBefore: 10,
				spaceAfter: 4
			}
		},
		{
			id: "Heading3",
			name: "Heading 3",
			rPr: {
				bold: !0,
				size: 12
			},
			pPr: {
				spaceBefore: 8,
				spaceAfter: 4
			}
		}
	],
	defaultFont: "Calibri",
	defaultSize: 11,
	format: "docx"
}), d = (e) => e.kind === "p", f = (e) => e.kind === "tbl";
function p(e) {
	return f(e) ? e.rows.map((e) => e.map((e) => e.paras.map(m).join(" ")).join("	")).join("\n") : m(e);
}
function m(e) {
	return e.runs.map((e) => e.text ?? (e.tab ? "	" : e.br ? "\n" : "")).join("");
}
function h(e) {
	return e.blocks.map(p).join("\n");
}
function g(e) {
	let t = h(e).trim();
	return t ? t.split(/\s+/).length : 0;
}
var ee = (e) => JSON.parse(JSON.stringify(e));
// src/engine/docx-read.ts
e.w;
var _ = e.r, v = 9525, y = 15, b = (e, t) => e.children.filter((e) => e.localName === t), x = (e, t) => e.children.find((e) => e.localName === t), S = (e, t) => {
	if (!e) return null;
	let n = e.getAttribute(t);
	if (n != null) return n;
	let r = e.attributes.find((e) => e.localName === t);
	return r ? r.value : null;
}, C = (e) => e != null && e !== "0" && e !== "false" && e !== "none";
function w(e) {
	if (!e) return;
	let t = {};
	x(e, "b") && (t.bold = C(S(x(e, "b"), "val") || "1")), x(e, "i") && (t.italic = C(S(x(e, "i"), "val") || "1"));
	let n = x(e, "u");
	if (n) {
		t.underline = !0;
		let e = S(n, "val");
		e === "double" ? t.underlineStyle = "double" : e === "words" && (t.underlineStyle = "words");
	}
	x(e, "strike") && (t.strike = C(S(x(e, "strike"), "val") || "1")), C(S(x(e, "caps"), "val") || "1") ? t.caps = "all" : C(S(x(e, "smallCaps"), "val") || "1") && (t.caps = "small"), C(S(x(e, "vanish"), "val") || "1") && (t.hidden = !0);
	let r = x(e, "rFonts");
	r && (t.font = S(r, "ascii") || S(r, "hAnsi") || void 0);
	let i = x(e, "sz");
	i && (t.size = (S(i, "val") || 22) / 2);
	let a = S(x(e, "color"), "val")?.toLowerCase();
	a && a !== "auto" && a !== "ffffff" && a !== "fff" && (t.color = "#" + a);
	let o = x(e, "highlight");
	o && (t.highlight = S(o, "val") || void 0);
	let s = x(e, "vertAlign");
	return s && (t.vertAlign = S(s, "val") || void 0), Object.keys(t).length ? t : void 0;
}
function T(e) {
	if (!e) return;
	let t = {};
	x(e, "pageBreakBefore") && (t.pageBreakBefore = !0);
	let n = x(e, "jc");
	if (n) {
		let e = S(n, "val") || "left";
		t.align = e === "both" ? "justify" : e === "start" ? "left" : e === "end" ? "right" : e;
	}
	let r = x(e, "pStyle");
	r && (t.style = S(r, "val") || void 0);
	let i = x(e, "numPr");
	if (i) {
		let e = x(i, "ilvl"), n = x(i, "numId");
		n && (t.numId = +(S(n, "val") || 0)), e && (t.ilvl = +(S(e, "val") || 0));
	}
	let a = x(e, "ind");
	if (a) {
		let e = S(a, "left") || S(a, "start");
		e && (t.indent = Math.round(+e / y));
		let n = S(a, "hanging");
		n && (t.hanging = Math.round(+n / y));
	}
	let o = x(e, "spacing");
	if (o) {
		let e = S(o, "before"), n = S(o, "after"), r = S(o, "line");
		e && (t.spaceBefore = Math.round(+e / y)), n && (t.spaceAfter = Math.round(+n / y)), r && S(o, "lineRule") !== "auto" ? t.line = Math.round(+r / y) : r && (t.line = r / 240);
	}
	x(e, "keepNext") && (t.keepNext = !0), x(e, "keepLines") && (t.keepLines = !0);
	let s = x(e, "outlineLvl");
	s && (t.outlineLvl = +(S(s, "val") || 9));
	let c = x(e, "pBdr");
	if (c) {
		let e = {};
		for (let t of [
			"top",
			"bottom",
			"left",
			"right",
			"between",
			"insideH",
			"insideV"
		]) {
			let n = x(c, t);
			if (!n) continue;
			let r = S(n, "val") || "single";
			r !== "nil" && r !== "none" && (e[t] = {
				val: r,
				sz: Math.round((S(n, "sz") || 4) / 8 * 100) / 100,
				color: (S(n, "color") || "auto") === "auto" ? "auto" : (S(n, "color") || "auto").toLowerCase(),
				space: +(S(n, "space") || 0)
			});
		}
		Object.keys(e).length && (t.borders = e);
	}
	let l = x(e, "shd");
	return l && (t.shading = "#" + (S(l, "fill") || "FFFFFF")), Object.keys(t).length ? t : void 0;
}
function E(e, t, n) {
	let r = w(x(e, "rPr"));
	if (x(e, "br")) return {
		br: !0,
		rPr: r
	};
	if (x(e, "tab")) return {
		tab: !0,
		rPr: r
	};
	let i = x(e, "drawing");
	if (i) {
		let a = te(i, t, n);
		return a ? {
			image: a,
			rPr: r
		} : {
			raw: D(e),
			rPr: r
		};
	}
	let a = x(e, "t");
	if (a) return {
		text: a.textContent,
		rPr: r
	};
	let o = e.textContent;
	return o ? {
		text: o,
		rPr: r
	} : {
		raw: D(e),
		rPr: r
	};
}
function D(e) {
	let t = e.attributes.map((e) => ` ${e.name}="${e.value.replace(/"/g, "&quot;")}"`).join(""), n = e.childNodes.map((e) => e.nodeType === 3 ? e.data.replace(/&/g, "&amp;").replace(/</g, "&lt;") : D(e)).join("");
	return `<${e.nodeName}${t}>${n}</${e.nodeName}>`;
}
function te(t, n, r) {
	let i = t.getElementsByTagNameNS("*", "extent")[0], a = t.getElementsByTagNameNS(e.a, "blip")[0];
	if (!a) return null;
	let o = a.getAttributeNS(_, "embed") || a.getAttribute("r:embed");
	if (!o) return null;
	let s = n.rel(r, o);
	if (!s) return null;
	let c = s.target.startsWith("/") ? s.target.slice(1) : ne(r, s.target), l = n.bytes(c);
	if (!l) return null;
	let u = c.split(".").pop() || "png", d = u === "jpg" || u === "jpeg" ? "image/jpeg" : u === "gif" ? "image/gif" : u === "svg" ? "image/svg+xml" : "image/" + u, f = URL.createObjectURL(new Blob([l], { type: d })), p = +(S(i, "cx") || 0), m = +(S(i, "cy") || 0);
	return {
		rId: o,
		part: c,
		url: f,
		w: Math.round(p / v) || 100,
		h: Math.round(m / v) || 100
	};
}
function ne(e, t) {
	let n = e.includes("/") ? e.slice(0, e.lastIndexOf("/") + 1) : "", r = [];
	for (let e of (n + t).split("/")) e !== "." && e !== "" && (e === ".." ? r.pop() : r.push(e));
	return r.join("/");
}
function O(e, t, n) {
	let r = c(T(x(e, "pPr")));
	for (let i of e.children) if (i.localName === "r") r.runs.push(E(i, t, n));
	else if (i.localName === "hyperlink") {
		let e = i.getAttributeNS(_, "id") || i.getAttribute("r:id") || "", a = (e ? t.rel(n, e) : void 0)?.target || "";
		for (let o of b(i, "r")) {
			let i = E(o, t, n);
			i.hyperlink = {
				rId: e,
				url: a
			}, r.runs.push(i);
		}
	} else i.localName === "pPr" || i.localName === "bookmarkStart" || i.localName === "bookmarkEnd" || i.localName === "proofErr" || r.runs.push({ raw: D(i) });
	return r;
}
function re(e, t, n) {
	let r = x(e, "tblGrid"), i = r ? b(r, "gridCol").map((e) => Math.round(+(S(e, "w") || 1e3) / y)) : [], a = [];
	for (let r of b(e, "tr")) {
		let e = [];
		for (let i of b(r, "tc")) {
			let r = x(i, "tcPr"), a = { paras: [] }, o = r ? x(r, "gridSpan") : void 0;
			o && (a.gridSpan = +(S(o, "val") || 1));
			let s = r ? x(r, "vMerge") : void 0;
			s && (a.vMerge = S(s, "val") === "restart" ? "restart" : "continue");
			let l = r ? x(r, "shd") : void 0;
			l && S(l, "fill") && S(l, "fill") !== "auto" && (a.fill = "#" + S(l, "fill"));
			let u = r ? x(r, "tcW") : void 0;
			u && (a.width = Math.round(+(S(u, "w") || 0) / y));
			let d = r ? x(r, "vAlign") : void 0;
			d && (a.vAlign = S(d, "val") === "center" ? "center" : S(d, "val") === "bottom" ? "bottom" : "top");
			let f = r ? x(r, "tcBorders") : void 0;
			if (f) {
				let e = {};
				for (let t of [
					"top",
					"bottom",
					"left",
					"right",
					"insideH",
					"insideV"
				]) {
					let n = x(f, t);
					if (!n) continue;
					let r = S(n, "val") || "single";
					r !== "nil" && r !== "none" && (e[t] = {
						val: r,
						sz: Math.round((S(n, "sz") || 4) / 8 * 100) / 100,
						color: (S(n, "color") || "auto") === "auto" ? "auto" : (S(n, "color") || "auto").toLowerCase(),
						space: +(S(n, "space") || 0)
					});
				}
				Object.keys(e).length && (a.borders = e);
			}
			let p = r ? x(r, "textDirection") : void 0;
			p && S(p, "val") === "tbRl" && (a.textDirection = "vertical");
			for (let e of b(i, "p")) a.paras.push(O(e, t, n));
			a.paras.length || a.paras.push(c()), e.push(a);
		}
		a.push(e);
	}
	let o = {
		kind: "tbl",
		id: s("t"),
		rows: a,
		widths: i
	}, l = x(e, "tblPr");
	if (l) {
		let t = x(l, "jc");
		t && (o.align = S(t, "val") === "center" ? "center" : S(t, "val") === "right" ? "right" : "left");
		let n = x(l, "tblLayout");
		n && S(n, "type") === "autofit" && (o.autoFit = "content");
		let r = b(e, "tr")[0], i = r ? x(r, "trPr") : void 0;
		i && x(i, "tblHeader") && (o.headerRows = 1);
	}
	return o;
}
function ie(e) {
	let t = [], n = "Calibri", r = 11;
	if (!e.has("word/styles.xml")) return {
		styles: t,
		defaultFont: n,
		defaultSize: r
	};
	let i = e.xml("word/styles.xml"), a = x(i, "docDefaults"), o = w(a ? x(x(a, "rPrDefault") || a, "rPr") : void 0);
	o?.font && (n = o.font), o?.size && (r = o.size);
	for (let e of b(i, "style")) {
		let n = S(e, "styleId") || "", r = {
			id: n,
			name: S(x(e, "name"), "val") || n
		}, i = x(e, "basedOn");
		i && (r.basedOn = S(i, "val") || void 0), r.rPr = w(x(e, "rPr")), r.pPr = T(x(e, "pPr")), t.push(r);
	}
	return {
		styles: t,
		defaultFont: n,
		defaultSize: r
	};
}
function ae(e) {
	let t = {};
	if (!e.has("word/numbering.xml")) return t;
	let n = e.xml("word/numbering.xml"), r = {};
	for (let e of b(n, "abstractNum")) {
		let t = S(e, "abstractNumId") || "", n = b(e, "lvl").find((e) => S(e, "ilvl") === "0") || b(e, "lvl")[0];
		r[t] = n && S(x(n, "numFmt"), "val") || "bullet";
	}
	for (let e of b(n, "num")) {
		let n = +(S(e, "numId") || 0);
		t[n] = r[S(x(e, "abstractNumId"), "val") || ""] || "bullet";
	}
	return t;
}
function oe(e) {
	return k(x(e, "sectPr"));
}
function k(e) {
	let t = l();
	if (!e) return t;
	let n = x(e, "pgSz");
	n && (t.pageW = +(S(n, "w") || t.pageW), t.pageH = +(S(n, "h") || t.pageH));
	let r = x(e, "pgMar");
	r && (t.margins = {
		top: +(S(r, "top") || 1440),
		right: +(S(r, "right") || 1440),
		bottom: +(S(r, "bottom") || 1440),
		left: +(S(r, "left") || 1440)
	});
	let i = x(e, "cols");
	return i && (t.columns = Math.max(1, +(S(i, "num") || 1))), t;
}
async function se(e) {
	let n = await t.open(e), r = n.main || "word/document.xml", i = n.xml(r), a = x(i, "body") || i, o = [], s = oe(a);
	for (let e of a.children) if (e.localName === "p") {
		o.push(O(e, n, r));
		let t = x(e, "pPr"), i = t ? x(t, "sectPr") : void 0;
		i && (s = k(i));
	} else e.localName === "tbl" ? o.push(re(e, n, r)) : e.localName === "sectPr" || o.push({
		...c(),
		raw: D(e)
	});
	o.length || o.push(c());
	let { styles: l, defaultFont: u, defaultSize: d } = ie(n);
	return {
		doc: {
			blocks: o,
			section: s,
			styles: l,
			defaultFont: u,
			defaultSize: d,
			numbering: ae(n),
			format: n.variant("docx")
		},
		pkg: n,
		part: r
	};
}
// src/engine/docx-write.ts
var ce = 15, A = 9525, j = (e) => Math.round(e * ce), M = "xmlns:w=\"" + e.w + "\" xmlns:r=\"" + e.r + "\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"" + e.a + "\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"";
function N(e) {
	if (!e) return "";
	let t = "";
	return e.font && (t += `<w:rFonts w:ascii="${n(e.font)}" w:hAnsi="${n(e.font)}" w:cs="${n(e.font)}"/>`), e.bold && (t += "<w:b/>"), e.italic && (t += "<w:i/>"), e.strike && (t += "<w:strike/>"), e.underline && (t += `<w:u w:val="${e.underlineStyle || "single"}"/>`), e.caps === "all" && (t += "<w:caps/>"), e.caps === "small" && (t += "<w:smallCaps/>"), e.hidden && (t += "<w:vanish/>"), e.color && (t += `<w:color w:val="${e.color.replace("#", "").toUpperCase()}"/>`), e.size && (t += `<w:sz w:val="${Math.round(e.size * 2)}"/><w:szCs w:val="${Math.round(e.size * 2)}"/>`), e.highlight && (t += `<w:highlight w:val="${n(e.highlight)}"/>`), e.vertAlign && (t += `<w:vertAlign w:val="${e.vertAlign}"/>`), t ? `<w:rPr>${t}</w:rPr>` : "";
}
function P(e) {
	if (!e) return "";
	let t = "";
	if (e.pageBreakBefore && (t += "<w:pageBreakBefore/>"), e.style && (t += `<w:pStyle w:val="${n(e.style)}"/>`), e.numId != null && (t += `<w:numPr><w:ilvl w:val="${e.ilvl || 0}"/><w:numId w:val="${e.numId}"/></w:numPr>`), e.spaceBefore != null || e.spaceAfter != null || e.line != null) {
		let n = e.spaceBefore == null ? "" : ` w:before="${j(e.spaceBefore)}"`, r = e.spaceAfter == null ? "" : ` w:after="${j(e.spaceAfter)}"`, i = "";
		e.line != null && (i = e.line > 4 ? ` w:line="${j(e.line)}" w:lineRule="exact"` : ` w:line="${Math.round(e.line * 240)}" w:lineRule="auto"`), t += `<w:spacing${n}${r}${i}/>`;
	}
	if (e.indent != null || e.hanging != null) {
		let n = e.indent == null ? "" : ` w:left="${j(e.indent)}"`, r = e.hanging == null ? "" : ` w:hanging="${j(e.hanging)}"`;
		t += `<w:ind${n}${r}/>`;
	}
	if (e.align && (t += `<w:jc w:val="${e.align === "justify" ? "both" : e.align}"/>`), e.keepNext && (t += "<w:keepNext/>"), e.keepLines && (t += "<w:keepLines/>"), e.outlineLvl != null && (t += `<w:outlineLvl w:val="${e.outlineLvl}"/>`), e.borders) {
		let n = Object.entries(e.borders).filter(([, e]) => e && e.val !== "nil" && e.val !== "none");
		n.length && (t += "<w:pBdr>" + n.map(([e, t]) => {
			let n = Math.max(2, Math.round((t?.sz ?? .5) * 8)), r = t?.val === "dashed" ? "dashed" : t?.val === "dotted" ? "dotted" : t?.val === "double" ? "double" : t?.val === "dotDash" ? "dotDash" : "single", i = (t?.color && t.color !== "auto" ? t.color : "auto").toUpperCase();
			return `<w:${e} w:val="${r}" w:sz="${n}" w:space="${t?.space ?? 1}" w:color="${i}"/>`;
		}).join("") + "</w:pBdr>");
	}
	return e.shading && (t += `<w:shd w:val="clear" w:color="auto" w:fill="${e.shading.replace("#", "").toUpperCase()}"/>`), t ? `<w:pPr>${t}</w:pPr>` : "";
}
function le(t, n, r) {
	if (r && t.pkg.rel(t.part, r)) return r;
	let i = t.pkg.relsOf(t.part).find((e) => e.target === n && e.mode === "External");
	if (i) return i.id;
	let a = n;
	if (t.linkIds.has(a)) return t.linkIds.get(a);
	let o = t.pkg.addRel(t.part, e.rel + "/hyperlink", n, "External");
	return t.linkIds.set(a, o), o;
}
async function ue(t, n) {
	if (n.part && t.pkg.has(n.part) && n.rId && t.pkg.rel(t.part, n.rId)) return n.rId;
	let r = await fetch(n.url), i = new Uint8Array(await r.arrayBuffer()), a = (n.part.split(".").pop() || "png").toLowerCase(), o = t.pkg.nextMediaName(a === "jpg" ? "jpeg" : a), s = (t.part.startsWith("word/") ? "word/media/" : "media/") + o;
	t.pkg.addPart(s, i);
	let c = t.pkg.addRel(t.part, e.rel + "/image", "media/" + o);
	return n.part = s, n.rId = c, c;
}
async function de(e, t) {
	if (t.raw != null) return t.raw;
	let n = N(t.rPr);
	if (t.image) {
		let r = await ue(e, t.image), i = Math.round(t.image.w * A), a = Math.round(t.image.h * A), o = ++e.picId;
		return `<w:r>${n}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${i}" cy="${a}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${o}" name="Picture ${o}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${o}" name="Picture ${o}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${r}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${i}" cy="${a}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
	}
	let i = "";
	if (t.br && (i += "<w:br/>"), t.tab && (i += "<w:tab/>"), t.text != null) {
		let e = /^\s|\s$/.test(t.text) ? " xml:space=\"preserve\"" : "";
		i += `<w:t${e}>${r(t.text)}</w:t>`;
	}
	i ||= "<w:t xml:space=\"preserve\"></w:t>";
	let a = `<w:r>${n}${i}</w:r>`;
	return t.hyperlink ? `<w:hyperlink r:id="${le(e, t.hyperlink.url, t.hyperlink.rId)}">${a}</w:hyperlink>` : a;
}
async function F(e, t) {
	if (t.raw != null && !t.runs.length) return t.raw;
	let n = [];
	for (let r of t.runs) n.push(await de(e, r));
	return `<w:p>${P(t.pPr)}${n.join("")}</w:p>`;
}
async function fe(e, t) {
	let n = t.widths.map((e) => `<w:gridCol w:w="${j(e)}"/>`).join(""), r = t.rows.some((e) => e.some((e) => !e.borders)) ? "" : "<w:tblBorders>" + [
		"top",
		"left",
		"bottom",
		"right",
		"insideH",
		"insideV"
	].map((e) => `<w:${e} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join("") + "</w:tblBorders>", i = t.align ? `<w:jc w:val="${t.align}"/>` : "", a = t.autoFit === "content" ? "<w:tblLayout w:type=\"autofit\"/>" : t.autoFit === "fixed" ? "<w:tblLayout w:type=\"fixed\"/>" : "", o = [];
	for (let n = 0; n < t.rows.length; n++) {
		let r = t.rows[n], i = [];
		for (let t of r) {
			let n = `<w:tcW w:w="${j(t.width || 100)}" w:type="dxa"/>`;
			if (t.gridSpan && t.gridSpan > 1 && (n += `<w:gridSpan w:val="${t.gridSpan}"/>`), t.vMerge && (n += t.vMerge === "restart" ? "<w:vMerge w:val=\"restart\"/>" : "<w:vMerge/>"), t.fill && (n += `<w:shd w:val="clear" w:color="auto" w:fill="${t.fill.replace("#", "").toUpperCase()}"/>`), t.vAlign && (n += `<w:vAlign w:val="${t.vAlign === "center" ? "center" : t.vAlign === "bottom" ? "bottom" : "top"}"/>`), t.textDirection === "vertical" && (n += "<w:textDirection w:val=\"tbRl\"/>"), t.borders && typeof t.borders == "object") {
				let e = Object.entries(t.borders).filter(([, e]) => e && e.val !== "nil" && e.val !== "none");
				e.length && (n += "<w:tcBorders>" + e.map(([e, t]) => {
					let n = Math.max(2, Math.round((t?.sz ?? .5) * 8));
					return `<w:${e} w:val="${t?.val === "dashed" ? "dashed" : t?.val === "dotted" ? "dotted" : t?.val === "double" ? "double" : "single"}" w:sz="${n}" w:space="0" w:color="${(t?.color && t.color !== "auto" ? t.color : "auto").toUpperCase()}"/>`;
				}).join("") + "</w:tcBorders>");
			}
			let r = [];
			for (let n of t.paras) r.push(await F(e, n));
			i.push(`<w:tc><w:tcPr>${n}</w:tcPr>${r.join("")}</w:tc>`);
		}
		let a = n === 0 && t.headerRows ? "<w:trPr><w:tblHeader/></w:trPr>" : "";
		o.push(`<w:tr>${a}${i.join("")}</w:tr>`);
	}
	return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${r}${i}${a}</w:tblPr><w:tblGrid>${n}</w:tblGrid>${o.join("")}</w:tbl>`;
}
function pe(e) {
	let t = e.section || l(), n = t.margins;
	return `<w:sectPr><w:pgSz w:w="${t.pageW}" w:h="${t.pageH}"/><w:pgMar w:top="${n.top}" w:right="${n.right}" w:bottom="${n.bottom}" w:left="${n.left}" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:num="${Math.max(1, t.columns || 1)}" w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>`;
}
async function me(e, t) {
	let n = [];
	for (let r of t.blocks) d(r) ? n.push(await F(e, r)) : f(r) && n.push(await fe(e, r));
	return n.join("") + pe(t);
}
var I = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", he = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
function L(e) {
	let t = e.styles.map((e) => {
		let t = `<w:name w:val="${n(e.name)}"/>`;
		return e.basedOn && (t += `<w:basedOn w:val="${n(e.basedOn)}"/>`), t += "<w:qFormat/>", e.pPr && (t += P(e.pPr).replace("<w:pPr>", "<w:pPr>").replace("</w:pPr>", "</w:pPr>")), e.rPr && (t += N(e.rPr)), `<w:style w:type="paragraph" w:styleId="${n(e.id)}">${t}</w:style>`;
	}).join("");
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:styles ${M}><w:docDefaults><w:rPrDefault>${`<w:rPr><w:rFonts w:ascii="${n(e.defaultFont)}" w:hAnsi="${n(e.defaultFont)}"/><w:sz w:val="${Math.round(e.defaultSize * 2)}"/></w:rPr>`}</w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` + t + "</w:styles>";
}
var R = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r
<w:settings ${M}><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/></w:settings>`;
async function ge(t, n = {}) {
	let r = i[n.format || t.format || "docx"] || i.docx, o = n.part || n.pkg?.main || "word/document.xml", s = n.pkg || a(o, "", r), c = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:document ${M}><w:body>${await me({
		pkg: s,
		part: o,
		picId: 100,
		linkIds: /* @__PURE__ */ new Map(),
		newImages: []
	}, t)}</w:body></w:document>`;
	return s.setPart(o, c), n.pkg ? s.has("word/styles.xml") || (s.addPart("word/styles.xml", L(t), I), s.addRel(o, e.rel + "/styles", "styles.xml")) : (s.addPart("word/styles.xml", L(t), I), s.addRel(o, e.rel + "/styles", "styles.xml"), s.addPart("word/settings.xml", R, he), s.addRel(o, e.rel + "/settings", "settings.xml")), s.save(r.mime);
}
// src/engine/render.ts
var _e = {
	"#ffff00": "yellow",
	"#00ff00": "green",
	"#00ffff": "cyan",
	"#ff00ff": "magenta",
	"#0000ff": "blue",
	"#ff0000": "red",
	"#000080": "darkBlue",
	"#008080": "darkCyan",
	"#008000": "darkGreen",
	"#800080": "darkMagenta",
	"#800000": "darkRed",
	"#808000": "darkYellow",
	"#808080": "darkGray",
	"#c0c0c0": "lightGray",
	"#000000": "black",
	"#ffffff": "white"
}, z = (e) => e.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"), ve = {
	arial: "\"Arimo\",\"Liberation Sans\",Helvetica,sans-serif",
	"arial black": "\"Arial Black\",\"Arimo\",sans-serif",
	"arial narrow": "\"Arial Narrow\",\"Arimo\",sans-serif",
	helvetica: "\"Helvetica Neue\",\"TeX Gyre Heros\",\"Nimbus Sans\",Arial,\"Arimo\",sans-serif",
	"times new roman": "\"Tinos\",\"Liberation Serif\",Times,serif",
	times: "\"Times New Roman\",\"Tinos\",\"TeX Gyre Termes\",serif",
	"courier new": "\"Cousine\",\"Liberation Mono\",Courier,monospace",
	courier: "\"Courier New\",\"Cousine\",monospace",
	calibri: "\"Carlito\",Candara,\"Segoe UI\",sans-serif",
	"calibri light": "\"Carlito\",sans-serif",
	aptos: "Calibri,\"Carlito\",sans-serif",
	georgia: "\"Gelasio\",\"DejaVu Serif\",\"Times New Roman\",\"Tinos\",serif",
	cambria: "\"Caladea\",Georgia,serif",
	tahoma: "\"DejaVu Sans Condensed\",Verdana,Arial,\"Arimo\",sans-serif",
	verdana: "\"DejaVu Sans\",Tahoma,Arial,\"Arimo\",sans-serif",
	symbol: "\"Segoe UI Symbol\",serif",
	wingdings: "\"Segoe UI Symbol\",sans-serif"
};
function ye(e) {
	let t = String(e).replace(/"/g, "").trim();
	return `'${t}',${(ve[t.toLowerCase()] || "Arial,\"Arimo\",sans-serif").replace(/"/g, "'")}`;
}
function B(e, t) {
	let n = [], r = e?.font || t.defaultFont, i = e?.size || t.defaultSize;
	n.push(`font-family:${ye(r)}`), n.push(`font-size:${i}pt`), e?.bold && n.push("font-weight:bold"), e?.italic && n.push("font-style:italic");
	let a = [];
	return e?.underline && a.push("underline"), e?.strike && a.push("line-through"), a.length && n.push("text-decoration-line:" + a.join(" ")), e?.underlineStyle === "double" && n.push("text-decoration-style:double"), e?.underlineStyle === "words" && n.push("text-decoration-style:dotted"), e?.caps === "all" && n.push("text-transform:uppercase"), e?.caps === "small" && n.push("font-variant:small-caps"), e?.color && !/^#(?:fff|ffffff)$/i.test(e.color.trim()) && n.push("color:" + e.color), e?.highlight && n.push("background:" + e.highlight), e?.vertAlign === "superscript" && n.push("vertical-align:super;font-size:smaller"), e?.vertAlign === "subscript" && n.push("vertical-align:sub;font-size:smaller"), n.join(";");
}
function be(e, t) {
	let n = [], r = e.pPr?.style ? t.styles.find((t) => t.id === e.pPr.style) : void 0, i = e.pPr?.align || r?.pPr?.align;
	i && n.push("text-align:" + i);
	let a = e.pPr?.spaceBefore ?? r?.pPr?.spaceBefore, o = e.pPr?.spaceAfter ?? r?.pPr?.spaceAfter;
	a != null && n.push("margin-top:" + a + "px"), o != null && n.push("margin-bottom:" + o + "px");
	let s = e.pPr?.indent ?? 0;
	s && n.push("padding-left:" + s + "px"), e.pPr?.line && n.push(e.pPr.line > 4 ? "line-height:" + e.pPr.line + "px" : "line-height:" + e.pPr.line), e.pPr?.pageBreakBefore && n.push("break-before:page;page-break-before:always"), e.pPr?.keepNext && n.push("page-break-after:avoid"), e.pPr?.keepLines && n.push("page-break-inside:avoid");
	let c = e.pPr?.borders;
	if (c) for (let e of [
		"top",
		"bottom",
		"left",
		"right"
	]) {
		let t = c[e];
		if (!t || t.val === "nil" || t.val === "none") continue;
		let r = Math.max(.5, (t.sz ?? .5) * 1.33), i = t.val === "dashed" || t.val === "dotted" || t.val === "dotDash" ? t.val === "dotted" ? "dotted" : "dashed" : t.val === "double" ? "double" : "solid";
		n.push(`border-${e}:${r}px ${i} ${t.color && t.color !== "auto" ? "#" + t.color : "#000"}`);
	}
	e.pPr?.shading && n.push("background:" + e.pPr.shading);
	let l = r?.rPr;
	return l && n.push(B(l, t)), n.join(";");
}
function xe(e, t) {
	if (e.raw != null) return `<span class="raw" contenteditable="false" data-raw="${encodeURIComponent(e.raw)}">[object]</span>`;
	let n = B(e.rPr, t);
	if (e.image) {
		let t = e.image, r = [n];
		t.rotation && r.push(`transform:rotate(${t.rotation}deg)`), t.contrast != null && r.push(`filter:contrast(${t.contrast}%)`), t.brightness != null && r.push(`filter:brightness(${t.brightness}%)`), t.gray && r.push("filter:grayscale(1)"), t.bw && r.push("filter:grayscale(1) contrast(1.6)"), t.washout && r.push("filter:brightness(1.4) contrast(0.6)"), t.offset && (t.offset.x || t.offset.y) && r.push(`position:relative;left:${t.offset.x || 0}px;top:${t.offset.y || 0}px`), (t.wrap === "square" || t.wrap === "tight") && r.push("float:left;margin:0 8px 4px 0"), t.wrap === "behind" && r.push("position:absolute;z-index:-1"), (t.wrap === "none" || t.wrap === "topBottom" || t.wrap === "through") && r.push("float:left;margin:0 8px 4px 0");
		let i = t.crop;
		return `<img class="img" src="${t.url}" width="${t.w}" height="${t.h}" data-part="${t.part}" data-rid="${t.rId}" data-w="${t.w}" data-h="${t.h}"${i ? ` data-crop="${i.l},${i.t},${i.r},${i.b}"` : ""} data-wrap="${t.wrap || "inline"}" style="${r.join(";")}">`;
	}
	if (e.br) return "<br>";
	if (e.tab) return `<span class="tb" style="${n}">\t</span>`;
	let r = `<span style="${n}">${z(e.text ?? "") || ""}</span>`;
	return e.hyperlink ? `<a href="${z(e.hyperlink.url)}" data-rid="${e.hyperlink.rId}" class="lnk">${r}</a>` : r;
}
function V(e, t) {
	if (d(e)) {
		let n = e.pPr?.style ? t.styles.find((t) => t.id === e.pPr.style) : void 0, r = e.runs.map((e) => xe({
			...e,
			rPr: {
				...n?.rPr,
				...e.rPr
			}
		}, t)).join(""), i = Se(e, t);
		return `<div class="p" data-id="${e.id}" style="${be(e, t)}">${i}${r || "<br>"}</div>`;
	}
	let n = e, r = n.widths.map((e) => `<col style="width:${e}px">`).join(""), i = [
		"border-collapse:collapse",
		`table-layout:${n.autoFit === "content" ? "auto" : "fixed"}`,
		"break-inside:avoid",
		"page-break-inside:avoid",
		n.align === "center" ? "margin-left:auto;margin-right:auto" : n.align === "right" ? "margin-left:auto" : ""
	].filter(Boolean).join(";"), a = n.rows.map((e, r) => `<tr>${e.map((e) => {
		let i = e.gridSpan && e.gridSpan > 1 ? ` colspan="${e.gridSpan}"` : "", a = e.vMerge === "continue" ? " data-vmerge=\"continue\"" : e.vMerge === "restart" ? " data-vmerge=\"restart\"" : "", o = Ce(e), s = [
			e.fill ? `background:${e.fill}` : "",
			o,
			"padding:2px 4px",
			`vertical-align:${e.vAlign === "center" ? "middle" : e.vAlign === "bottom" ? "bottom" : "top"}`,
			e.textDirection === "vertical" ? "writing-mode:vertical-lr" : "",
			r < (n.headerRows || 0) ? "font-weight:bold" : ""
		].filter(Boolean).join(";"), c = e.paras.map((e) => V(e, t)).join("");
		return `<td${i}${a} data-w="${e.width || 0}" style="${s}">${c}</td>`;
	}).join("")}</tr>`).join("");
	return `<table class="tbl" data-id="${n.id}" style="${i}"><colgroup>${r}</colgroup>${a}</table>`;
}
function Se(e, t) {
	return e.pPr?.numId == null ? "" : `<span class="lm" contenteditable="false" data-num="${e.pPr.numId}" data-ilvl="${e.pPr.ilvl || 0}"></span>`;
}
function Ce(e) {
	if (!e.borders) return "border:1px solid #888";
	let t = [];
	for (let n of [
		"top",
		"bottom",
		"left",
		"right"
	]) {
		let r = e.borders[n];
		if (!r || r.val === "nil" || r.val === "none") {
			t.push(`border-${n}:0`);
			continue;
		}
		let i = Math.max(.5, (r.sz ?? .5) * 1.33), a = r.val === "dashed" || r.val === "dotted" || r.val === "dotDash" ? r.val === "dotted" ? "dotted" : "dashed" : r.val === "double" ? "double" : "solid";
		t.push(`border-${n}:${i}px ${a} ${r.color && r.color !== "auto" ? "#" + r.color : "#000"}`);
	}
	return t.join(";");
}
function H(e) {
	if (!e) return;
	if (e.startsWith("#")) return e.length === 7 ? e : void 0;
	let t = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(e);
	if (!t) return;
	let n = (e) => (+e).toString(16).padStart(2, "0");
	return "#" + n(t[1]) + n(t[2]) + n(t[3]);
}
function we(e, t) {
	let n = { ...e }, r = t.tagName.toLowerCase();
	(r === "b" || r === "strong") && (n.bold = !0), (r === "i" || r === "em") && (n.italic = !0), r === "u" && (n.underline = !0), (r === "s" || r === "strike" || r === "del") && (n.strike = !0);
	let i = t.style;
	i.fontWeight && (i.fontWeight === "bold" || +i.fontWeight >= 600) && (n.bold = !0), i.fontWeight && (i.fontWeight === "normal" || +i.fontWeight < 600) && (n.bold = !1), i.fontStyle === "italic" && (n.italic = !0), i.fontStyle === "normal" && (n.italic = !1);
	let a = i.textDecorationLine || i.textDecoration || "";
	if (a.includes("underline") && (n.underline = !0), a.includes("line-through") && (n.strike = !0), i.fontFamily && (n.font = i.fontFamily.split(",")[0].replace(/["']/g, "").trim()), i.fontSize) {
		let e = parseFloat(i.fontSize);
		i.fontSize.endsWith("pt") ? n.size = e : isNaN(e) || (n.size = Math.round(e * .75 * 2) / 2);
	}
	let o = H(i.color);
	o && (n.color = o);
	let s = H(i.backgroundColor);
	return s && s !== "#ffffff" && (n.highlight = _e[s.toLowerCase()] || s), i.verticalAlign === "super" && (n.vertAlign = "superscript"), i.verticalAlign === "sub" && (n.vertAlign = "subscript"), i.textTransform === "uppercase" && (n.caps = "all"), i.fontVariantCaps === "small-caps" && (n.caps = "small"), i.textDecorationStyle === "double" && (n.underlineStyle = "double"), i.textDecorationStyle === "dotted" && (n.underlineStyle = "words"), (i.color === "rgba(0, 0, 0, 0)" || i.color === "transparent") && (n.hidden = !0), n;
}
function U(e, t, n, r) {
	if (e.nodeType === 3) {
		let i = e.nodeValue || "";
		i && r.push({
			text: i,
			rPr: W(t),
			hyperlink: n
		});
		return;
	}
	if (e.nodeType !== 1) return;
	let i = e, a = i.tagName.toLowerCase();
	if (a === "br") {
		r.push({
			br: !0,
			rPr: W(t)
		});
		return;
	}
	if (a === "img") {
		let e = (i.getAttribute("data-crop") || "").split(",").map(Number);
		r.push({
			image: {
				url: i.getAttribute("src") || "",
				part: i.getAttribute("data-part") || "",
				rId: i.getAttribute("data-rid") || "",
				w: +(i.getAttribute("data-w") || i.getAttribute("width") || 100),
				h: +(i.getAttribute("data-h") || i.getAttribute("height") || 100),
				...e.length === 4 && e.some((e) => e > 0) ? { crop: {
					l: e[0],
					r: e[2],
					t: e[1],
					b: e[3]
				} } : {},
				wrap: i.getAttribute("data-wrap") || "inline"
			},
			rPr: W(t)
		});
		return;
	}
	if (i.classList.contains("lm")) return;
	if (i.classList.contains("raw")) {
		r.push({ raw: decodeURIComponent(i.getAttribute("data-raw") || "") });
		return;
	}
	if (a === "span" && i.classList.contains("tb")) {
		r.push({
			tab: !0,
			rPr: W(t)
		});
		return;
	}
	let o = t, s = n;
	a === "a" ? (s = {
		rId: i.getAttribute("data-rid") || "",
		url: i.getAttribute("href") || ""
	}, o = {
		...t,
		color: t.color || "#0563c1",
		underline: !0
	}) : o = we(t, i);
	for (let e of Array.from(i.childNodes)) U(e, o, s, r);
}
function W(e) {
	let t = {};
	return e.bold && (t.bold = !0), e.italic && (t.italic = !0), e.underline && (t.underline = !0), e.strike && (t.strike = !0), e.font && (t.font = e.font), e.size && (t.size = e.size), e.color && (t.color = e.color), e.highlight && (t.highlight = e.highlight), e.vertAlign && (t.vertAlign = e.vertAlign), e.caps && (t.caps = e.caps), e.hidden && (t.hidden = e.hidden), e.underlineStyle && (t.underlineStyle = e.underlineStyle), Object.keys(t).length ? t : void 0;
}
function G(e, t) {
	let n = [];
	for (let t of Array.from(e.childNodes)) U(t, {}, void 0, n);
	let r = t ? {
		...t,
		runs: n
	} : {
		...c(),
		runs: n
	};
	r.id = e.getAttribute("data-id") || r.id;
	let i = e.style.textAlign;
	return i && (r.pPr = {
		...r.pPr,
		align: i
	}), r;
}
function Te(e, t, n) {
	let r = e.getAttribute("data-id") || void 0;
	if (e.tagName.toLowerCase() === "table") {
		let t = n && n.kind === "tbl" ? n : void 0, i = /* @__PURE__ */ new Map();
		for (let e of t?.rows || []) for (let t of e) for (let e of t.paras) i.set(e.id, e);
		let a = Array.from(e.querySelectorAll("col")).map((e) => parseInt(e.style.width) || 100), o = Array.from(e.querySelectorAll(":scope > tbody > tr, :scope > tr")).map((e) => Array.from(e.querySelectorAll(":scope > td, :scope > th")).map((e) => {
			let t = e, n = Array.from(t.querySelectorAll(":scope > .p")).map((e) => {
				let t = e;
				return G(t, i.get(t.getAttribute("data-id") || ""));
			});
			n.length || n.push(c());
			let r = { paras: n }, a = t.getAttribute("colspan");
			a && +a > 1 && (r.gridSpan = +a);
			let o = t.getAttribute("data-vmerge");
			(o === "restart" || o === "continue") && (r.vMerge = o);
			let s = H(t.style.backgroundColor);
			s && (r.fill = s);
			let l = +(t.getAttribute("data-w") || 0);
			l && (r.width = l);
			let u = t.style.verticalAlign;
			(u === "middle" || u === "bottom") && (r.vAlign = u === "middle" ? "center" : "bottom"), t.style.writingMode.includes("vertical") && (r.textDirection = "vertical");
			{
				let e = {};
				for (let n of [
					"top",
					"bottom",
					"left",
					"right"
				]) {
					let r = `border${n[0].toUpperCase()}${n.slice(1)}`, i = t.style, a = parseFloat(i[r + "Width"] || "0");
					if (!a) continue;
					let o = i[r + "Style"] || "solid", s = H(i[r + "Color"]) || "#000000";
					e[n] = {
						val: o === "none" ? "nil" : o,
						sz: Math.round(a / 1.33 * 100) / 100,
						color: s.replace("#", "")
					};
				}
				Object.keys(e).length && (r.borders = e);
			}
			return r;
		}));
		return {
			kind: "tbl",
			id: r || s("t"),
			rows: o,
			widths: a.length ? a : t?.widths || [100],
			...t?.align ? { align: t.align } : {},
			...t?.autoFit ? { autoFit: t.autoFit } : {},
			...t?.headerRows ? { headerRows: t.headerRows } : {}
		};
	}
	return G(e, n && n.kind === "p" ? n : void 0);
}
// src/engine/layout.ts
function Ee(e, t, n, r) {
	r.style.width = t + "px", r.innerHTML = "";
	let i = [[]], a = 0;
	for (let t of e.blocks) {
		let o = document.createElement("div");
		o.innerHTML = V(t, e);
		let s = o.firstElementChild;
		r.appendChild(s);
		let c = s.offsetHeight;
		r.removeChild(s);
		let l = t.kind === "p" && t.pPr?.pageBreakBefore;
		(a + c > n || l) && i[i.length - 1].length ? (i.push([t]), a = c) : (i[i.length - 1].push(t), a += c);
	}
	return i;
}
function De(e, t) {
	let n = /* @__PURE__ */ new Map();
	for (let r of Array.from(e.querySelectorAll(".lm"))) {
		let e = r.getAttribute("data-num") || "0", i = r.getAttribute("data-ilvl") || "0", a = t.numbering?.[+e] || "bullet", o = e + ":" + i;
		if (a === "bullet") r.textContent = [
			"•",
			"◦",
			"▪"
		][i % 3] + " ";
		else {
			let e = (n.get(o) || 0) + 1;
			n.set(o, e), r.textContent = (a === "lowerLetter" ? String.fromCharCode(96 + e) : a === "upperLetter" ? String.fromCharCode(64 + e) : a === "lowerRoman" ? Oe(e) : String(e)) + ". ";
		}
	}
}
function Oe(e) {
	let t = [
		[10, "x"],
		[9, "ix"],
		[5, "v"],
		[4, "iv"],
		[1, "i"]
	], n = "";
	for (let [r, i] of t) for (; e >= r;) n += i, e -= r;
	return n;
}
function ke(e) {
	let t = window.getSelection();
	if (!t || !t.rangeCount) return null;
	let n = t.getRangeAt(0), r = n.startContainer, i = null;
	for (; r && r !== e;) {
		if (r.nodeType === 1 && r.hasAttribute?.("data-id")) {
			i = r;
			break;
		}
		r = r.parentNode;
	}
	if (!i) return null;
	let a = n.cloneRange();
	a.selectNodeContents(i), a.setEnd(n.startContainer, n.startOffset);
	let o = a.toString().length;
	return {
		id: i.getAttribute("data-id"),
		offset: o
	};
}
function Ae(e, t) {
	if (!t) return;
	let n = e.querySelector(`[data-id="${t.id}"]`);
	if (!n) return;
	let r = document.createTreeWalker(n, NodeFilter.SHOW_TEXT), i = t.offset, a = r.nextNode(), o = null;
	for (; a;) {
		let e = (a.nodeValue || "").length;
		if (i <= e) {
			let e = window.getSelection(), t = document.createRange();
			t.setStart(a, i), t.collapse(!0), e.removeAllRanges(), e.addRange(t);
			return;
		}
		i -= e, o = a, a = r.nextNode();
	}
	let s = window.getSelection(), c = document.createRange();
	o ? c.setStart(o, (o.nodeValue || "").length) : c.selectNodeContents(n), c.collapse(!1), s.removeAllRanges(), s.addRange(c);
}
// src/io/markdown.ts
function je(e) {
	return {
		source: e,
		text: e.replace(/^\uFEFF/, "").replace(/\r\n?/g, "\n")
	};
}
function Me(e) {
	return e.replace(/\r\n?/g, "\n").replace(/\n?$/, "\n");
}
// src/io/html.ts
function Ne(e) {
	if (typeof DOMParser > "u") return e.replace(/<[^>]+>/g, "");
	let t = new DOMParser().parseFromString(e, "text/html");
	return t.body?.innerText ?? t.body?.textContent ?? "";
}
function Pe(e, t = "Document") {
	let n = (e) => e.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"), r = e.split(/\r?\n/).map((e) => `<p>${n(e) || "&nbsp;"}</p>`).join("\n");
	return `<!doctype html><html lang="en"><head><meta charset="utf-8"><title>${n(t)}</title></head><body>${r}</body></html>`;
}
// src/services/csv.ts
function Fe(e) {
	let t = e.replace(/^\uFEFF/, ""), n = t.split(/\r?\n/)[0] ?? "", r = [
		"	",
		";",
		","
	].sort((e, t) => n.split(t).length - 1 - (n.split(e).length - 1))[0], i = [], a = [], o = "", s = !1;
	for (let e = 0; e < t.length; e++) {
		let n = t[e];
		s ? n === "\"" && t[e + 1] === "\"" ? (o += "\"", e++) : n === "\"" ? s = !1 : o += n : n === "\"" && !o ? s = !0 : n === r ? (a.push(o), o = "") : n === "\n" || n === "\r" ? (n === "\r" && t[e + 1] === "\n" && e++, a.push(o), a.some(Boolean) && i.push(a), a = [], o = "") : o += n;
	}
	a.push(o), a.some(Boolean) && i.push(a);
	let c = (i.shift() ?? []).map((e, t) => e.trim() || `Field${t + 1}`);
	return {
		fields: c,
		rows: i.map((e) => c.map((t, n) => e[n] ?? ""))
	};
}
function Ie(e) {
	return /[",\r\n]/.test(e) ? `"${e.replace(/"/g, "\"\"")}"` : e;
}
function Le(e, t = ",") {
	return [e.fields, ...e.rows].map((e) => e.map(Ie).join(t)).join("\r\n") + "\r\n";
}
// src/services/autocorrect.ts
var Re = {
	"(c)": "©",
	"(r)": "®",
	"(tm)": "™",
	"...": "…",
	":-)": "🙂",
	":-(": "🙁",
	"-->": "→",
	"<--": "←",
	"1/2": "½",
	"1/4": "¼",
	"3/4": "¾"
};
function K(e, t = Re) {
	let n = t[e] ?? t[e.toLowerCase()];
	return n ? /^\p{Lu}/u.test(e) ? n[0].toUpperCase() + n.slice(1) : e === e.toUpperCase() && e.length > 1 ? n.toUpperCase() : n : e;
}
function ze(e, t) {
	return e.replace(/\S+/g, (e) => K(e, t));
}
// src/services/tables.ts
function q(e, t) {
	let n = Math.max(1, Math.floor(e)), r = Math.max(1, Math.floor(t));
	return {
		kind: "tbl",
		id: s("tbl"),
		widths: Array(r).fill(96),
		rows: Array.from({ length: n }, () => Array.from({ length: r }, () => ({ paras: [c()] })))
	};
}
function Be(e, t, n) {
	return e.rows[t]?.[n];
}
function Ve(e) {
	let t = e.widths.length || e.rows[0]?.length || 1;
	return e.rows.push(Array.from({ length: t }, () => ({ paras: [c()] }))), e;
}
function He(e) {
	e.widths.push(96);
	for (let t of e.rows) t.push({ paras: [c()] });
	return e;
}
function Ue(e, t) {
	return e.rows.length > 1 && e.rows.splice(t, 1), e;
}
function We(e, t) {
	if (e.widths.length > 1) {
		e.widths.splice(t, 1);
		for (let n of e.rows) n.splice(t, 1);
	}
	return e;
}
function Ge(e, t, n) {
	let r = q(t, n), i = e.blocks.length && e.blocks[e.blocks.length - 1].kind === "p" ? e.blocks.length - 1 : e.blocks.length;
	return e.blocks.splice(i, 0, r), r;
}
// src/services/fields.ts
function J(e) {
	let t = e.trim().match(/"(?:[^"]|"")*"|\S+/g) ?? [], n = (t.shift() ?? "").toUpperCase(), r = [], i = {};
	for (let e = 0; e < t.length; e++) {
		let n = t[e];
		n.startsWith("\\") ? i[n.slice(1).toLowerCase()] = t[e + 1]?.startsWith("\\") ? !0 : t[++e] ?? !0 : r.push(n.replace(/^"|"$/g, ""));
	}
	return {
		name: n,
		args: r,
		switches: i
	};
}
function Ke(e, t = {}) {
	let n = J(e), r = t.now ?? /* @__PURE__ */ new Date();
	switch (n.name) {
		case "PAGE": return String(t.page ?? 1);
		case "NUMPAGES":
		case "SECTIONPAGES": return String(t.pages ?? 1);
		case "TITLE": return t.title ?? "";
		case "AUTHOR": return t.author ?? "";
		case "FILENAME": return t.fileName ?? "";
		case "DATE":
		case "TIME": return r.toLocaleString();
		case "MERGEFIELD": return t[n.args[0]] ?? "";
		default: return n.args.join(" ");
	}
}
// src/services/spell.ts
function Y(e) {
	let t = [], n = /[\p{L}\p{M}\p{N}]+(?:['’\-][\p{L}\p{M}\p{N}]+)*/gu, r;
	for (; r = n.exec(e);) t.push({
		word: r[0],
		start: r.index,
		end: r.index + r[0].length
	});
	return t;
}
function qe(e) {
	let t = [];
	for (let n of e.matchAll(/\b(\w+)\s+\1\b/giu)) t.push({
		word: n[0],
		start: n.index ?? 0,
		end: (n.index ?? 0) + n[0].length,
		kind: "grammar",
		suggestions: [n[1]],
		message: "Repeated word"
	});
	return t;
}
function Je(e, t) {
	return Y(e).filter((e) => !t.has(e.word.toLocaleLowerCase())).map((e) => ({
		word: e.word,
		start: e.start,
		end: e.end,
		kind: "spelling",
		suggestions: [],
		message: "Not in dictionary"
	}));
}
function X(e, t) {
	let n = Array.from({ length: t.length + 1 }, (e, t) => t);
	for (let r = 1; r <= e.length; r++) {
		let i = n[0];
		n[0] = r;
		for (let a = 1; a <= t.length; a++) {
			let o = n[a];
			n[a] = Math.min(n[a] + 1, n[a - 1] + 1, i + (e[r - 1] === t[a - 1] ? 0 : 1)), i = o;
		}
	}
	return n[t.length];
}
function Ye(e, t, n = 5) {
	return Array.from(t).map((t) => [t, X(e.toLowerCase(), t.toLowerCase())]).sort((e, t) => e[1] - t[1]).slice(0, n).map((e) => e[0]);
}
// src/services/drawing.ts
function Xe(e, t = "") {
	return {
		image: { ...e },
		rPr: {},
		text: t || void 0
	};
}
function Ze(e, t, n = 640) {
	if (e <= n) return {
		width: e,
		height: t
	};
	let r = n / e;
	return {
		width: n,
		height: Math.max(1, Math.round(t * r))
	};
}
function Qe(e) {
	return {
		png: "image/png",
		jpg: "image/jpeg",
		jpeg: "image/jpeg",
		gif: "image/gif",
		webp: "image/webp",
		svg: "image/svg+xml"
	}[e.split(".").pop()?.toLowerCase() ?? ""] ?? "application/octet-stream";
}
// src/services/rtf.ts
function $e(e) {
	return e.replace(/\\'[0-9a-f]{2}/gi, (e) => String.fromCharCode(parseInt(e.slice(2), 16))).replace(/\\par[d]?\s?/gi, "\n").replace(/\\tab\s?/gi, "	").replace(/\\u(-?\d+)\??/g, (e, t) => String.fromCharCode(Number(t) < 0 ? Number(t) + 65536 : Number(t))).replace(/\\[a-z]+\d*\s?/gi, "").replace(/[{}]/g, "").replace(/\r\n?/g, "\n").trim();
}
function et(e) {
	return `{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 Calibri;}}\\f0\\fs22 ${e.replace(/\\/g, "\\\\").replace(/[{}]/g, (e) => `\\${e}`).replace(/\n/g, "\\par\n").replace(/\t/g, "\\tab ")}}`;
}
// src/services/review.ts
function tt(e, t) {
	return {
		...e,
		raw: e.raw,
		rPr: { ...e.rPr },
		text: t.kind === "delete" ? void 0 : e.text
	};
}
function nt(e) {
	return {
		...e,
		accepted: !0
	};
}
function rt(e) {
	return {
		...e,
		accepted: !1
	};
}
function it(e) {
	let t = [];
	for (let n of e.blocks) if (n.kind === "p") for (let e of n.runs) e.raw?.includes("comment") && t.push({
		id: e.raw,
		author: "",
		date: "",
		text: e.text ?? ""
	});
	return t;
}
function at(e) {
	return {
		kind: "p",
		id: `review-${Date.now()}`,
		runs: e ? [{ text: e }] : []
	};
}
// src/services/preserve.ts
function ot(e, t, n) {
	return e.has(t) || e.set(t, n), e;
}
function st(e) {
	return Array.from(e, (e) => ({
		...e,
		bytes: e.bytes.slice()
	}));
}
// src/services/commands.ts
var ct = class {
	commands = /* @__PURE__ */ new Map();
	register(e) {
		this.commands.set(e.id, e);
	}
	get(e) {
		return this.commands.get(e);
	}
	list() {
		return Array.from(this.commands.values());
	}
	async execute(e, t = {}) {
		let n = this.get(e);
		return !n || n.enabled?.(t) === !1 ? !1 : (await n.run(t), !0);
	}
}, lt = class {
	limit;
	past = [];
	future = [];
	constructor(e = 100) {
		this.limit = e;
	}
	push(e) {
		this.past.push(e), this.past.length > this.limit && this.past.shift(), this.future = [];
	}
	undo(e) {
		let t = this.past.pop();
		return t !== void 0 && this.future.push(e), t;
	}
	redo(e) {
		let t = this.future.pop();
		return t !== void 0 && this.past.push(e), t;
	}
	clear() {
		this.past = [], this.future = [];
	}
};
// src/services/layout.ts
function Z(e, t, n = 720) {
	let r = [], i = {
		page: 1,
		blocks: [],
		height: 0
	};
	for (let a of e.blocks) {
		let e = t.get(a.id) ?? 24;
		i.blocks.length && i.height + e > n && (r.push(i), i = {
			page: i.page + 1,
			blocks: [],
			height: 0
		}), i.blocks.push(a), i.height += e;
	}
	return (i.blocks.length || !r.length) && r.push(i), r;
}
function ut(e, t, n) {
	return Z(e, t, n).length;
}
// src/services/outline.ts
function dt(e) {
	let t = [], n = [];
	for (let r of e.blocks) {
		if (r.kind !== "p") continue;
		let e = Q(r);
		if (!e) continue;
		let i = {
			id: r.id,
			title: r.runs.map((e) => e.text ?? "").join("").trim() || "(blank)",
			level: e,
			block: r,
			children: []
		};
		for (; n.length && n[n.length - 1].level >= e;) n.pop();
		(n[n.length - 1]?.children ?? t).push(i), n.push(i);
	}
	return t;
}
function Q(e) {
	if (e.kind !== "p") return 0;
	let t = e.pPr?.style ?? "", n = /^Heading([1-9])$/i.exec(t);
	return n ? Number(n[1]) : 0;
}
// src/services/textOps.ts
function ft(e, t, n, r) {
	return e.slice(0, t) + r + e.slice(n);
}
function $(e, t) {
	return t === "upper" ? e.toUpperCase() : t === "lower" ? e.toLowerCase() : e.toLowerCase().replace(/(^|\s)(\p{L})/gu, (e, t, n) => t + n.toUpperCase());
}
function pt(e, t) {
	let n = e.runs, r = "", i = "", a = 0;
	for (let e of n) {
		let n = e.text ?? "", o = Math.max(0, Math.min(n.length, t - a));
		r += n.slice(0, o), i += n.slice(o), a += n.length;
	}
	return [{
		...e,
		id: `${e.id}-a`,
		runs: r ? [{ text: r }] : []
	}, {
		...e,
		id: `${e.id}-b`,
		runs: i ? [{ text: i }] : []
	}];
}
function mt(e, t) {
	return {
		text: e,
		rPr: { ...t }
	};
}
// src/services/omml.ts
function ht(e) {
	return e.replace(/<m:br\s*\/?>(\s*)/g, "\n").replace(/<[^>]+>/g, "").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&").replace(/\s+/g, " ").trim();
}
function gt(e) {
	return `<m:oMathPara><m:oMath><m:r><m:t>${e.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</m:t></m:r></m:oMath></m:oMathPara>`;
}
// src/services/templates.ts
var _t = class {
	templates;
	constructor(e = []) {
		this.templates = e;
	}
	list() {
		return this.templates.slice();
	}
	get(e) {
		return this.templates.find((t) => t.id === e);
	}
	create(e) {
		return this.get(e)?.content ?? "";
	}
};
// src/services/quireCommands.ts
function vt(e) {
	return [
		{
			label: "File",
			items: [
				{
					label: "New",
					key: "Ctrl+N",
					run: e.onNew
				},
				{
					label: "Open…",
					key: "Ctrl+O",
					run: e.onOpen
				},
				{
					label: "Save",
					key: "Ctrl+S",
					run: e.onSave
				}
			]
		},
		{
			label: "Edit",
			items: [
				{
					label: "Undo",
					key: "Ctrl+Z",
					run: e.onUndo
				},
				{
					label: "Redo",
					key: "Ctrl+Y",
					run: e.onRedo
				},
				{
					label: "Find…",
					key: "Ctrl+F",
					run: e.onFind
				}
			]
		},
		{
			label: "Insert",
			items: [{
				label: "Table…",
				run: e.onInsertTable
			}]
		}
	];
}
function yt(e) {
	return [
		{
			id: "new",
			label: "New",
			title: "New",
			run: e.onNew
		},
		{
			id: "open",
			label: "Open",
			title: "Open",
			run: e.onOpen
		},
		{
			id: "save",
			label: "Save",
			title: "Save",
			run: e.onSave
		},
		{
			id: "undo",
			label: "Undo",
			title: "Undo",
			run: e.onUndo
		},
		{
			id: "redo",
			label: "Redo",
			title: "Redo",
			run: e.onRedo
		},
		{
			id: "find",
			label: "Find",
			title: "Find",
			run: e.onFind
		}
	];
}
export { Ae as $, Je as A, Ue as B, $e as C, Qe as D, Ze as E, He as F, Le as G, ze as H, Ve as I, je as J, Ne as K, q as L, Y as M, Ke as N, qe as O, J as P, De as Q, Ge as R, rt as S, Xe as T, K as U, Be as V, Fe as W, ke as X, Me as Y, Ee as Z, ot as _, ht as a, ee as at, it as b, pt as c, c as ct, dt as d, u as dt, V as et, ut as f, m as ft, st as g, lt as h, gt as i, p as it, Ye as j, X as k, $ as l, d as lt, ct as m, g as mt, yt as n, ge as nt, mt as o, l as ot, Z as p, s as pt, Pe as q, _t as r, se as rt, ft as s, h as st, vt as t, Te as tt, Q as u, f as ut, nt as v, et as w, at as x, tt as y, We as z };
