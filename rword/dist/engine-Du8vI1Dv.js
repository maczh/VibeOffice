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
function ee(e) {
	let t = h(e).trim();
	return t ? t.split(/\s+/).length : 0;
}
var te = (e) => JSON.parse(JSON.stringify(e));
// src/engine/docx-read.ts
e.w;
var g = e.r, _ = 9525, v = 15, y = (e, t) => e.children.filter((e) => e.localName === t), b = (e, t) => e.children.find((e) => e.localName === t), x = (e, t) => {
	if (!e) return null;
	let n = e.getAttribute(t);
	if (n != null) return n;
	let r = e.attributes.find((e) => e.localName === t);
	return r ? r.value : null;
}, S = (e) => e != null && e !== "0" && e !== "false" && e !== "none";
function C(e) {
	if (!e) return;
	let t = {};
	b(e, "b") && (t.bold = S(x(b(e, "b"), "val") || "1")), b(e, "i") && (t.italic = S(x(b(e, "i"), "val") || "1"));
	let n = b(e, "u");
	if (n) {
		t.underline = !0;
		let e = x(n, "val");
		e === "double" ? t.underlineStyle = "double" : e === "words" && (t.underlineStyle = "words");
	}
	b(e, "strike") && (t.strike = S(x(b(e, "strike"), "val") || "1")), S(x(b(e, "caps"), "val") || "1") ? t.caps = "all" : S(x(b(e, "smallCaps"), "val") || "1") && (t.caps = "small"), S(x(b(e, "vanish"), "val") || "1") && (t.hidden = !0);
	let r = b(e, "rFonts");
	r && (t.font = x(r, "ascii") || x(r, "hAnsi") || void 0);
	let i = b(e, "sz");
	i && (t.size = (x(i, "val") || 22) / 2);
	let a = b(e, "color");
	a && x(a, "val") && x(a, "val") !== "auto" && (t.color = ("#" + x(a, "val")).toLowerCase());
	let o = b(e, "highlight");
	o && (t.highlight = x(o, "val") || void 0);
	let s = b(e, "vertAlign");
	return s && (t.vertAlign = x(s, "val") || void 0), Object.keys(t).length ? t : void 0;
}
function w(e) {
	if (!e) return;
	let t = {};
	b(e, "pageBreakBefore") && (t.pageBreakBefore = !0);
	let n = b(e, "jc");
	if (n) {
		let e = x(n, "val") || "left";
		t.align = e === "both" ? "justify" : e === "start" ? "left" : e === "end" ? "right" : e;
	}
	let r = b(e, "pStyle");
	r && (t.style = x(r, "val") || void 0);
	let i = b(e, "numPr");
	if (i) {
		let e = b(i, "ilvl"), n = b(i, "numId");
		n && (t.numId = +(x(n, "val") || 0)), e && (t.ilvl = +(x(e, "val") || 0));
	}
	let a = b(e, "ind");
	if (a) {
		let e = x(a, "left") || x(a, "start");
		e && (t.indent = Math.round(+e / v));
		let n = x(a, "hanging");
		n && (t.hanging = Math.round(+n / v));
	}
	let o = b(e, "spacing");
	if (o) {
		let e = x(o, "before"), n = x(o, "after"), r = x(o, "line");
		e && (t.spaceBefore = Math.round(+e / v)), n && (t.spaceAfter = Math.round(+n / v)), r && x(o, "lineRule") !== "auto" ? t.line = Math.round(+r / v) : r && (t.line = r / 240);
	}
	b(e, "keepNext") && (t.keepNext = !0), b(e, "keepLines") && (t.keepLines = !0);
	let s = b(e, "outlineLvl");
	s && (t.outlineLvl = +(x(s, "val") || 9)), b(e, "pBdr") && (t.borders = !0);
	let c = b(e, "shd");
	return c && (t.shading = "#" + (x(c, "fill") || "FFFFFF")), Object.keys(t).length ? t : void 0;
}
function T(e, t, n) {
	let r = C(b(e, "rPr"));
	if (b(e, "br")) return {
		br: !0,
		rPr: r
	};
	if (b(e, "tab")) return {
		tab: !0,
		rPr: r
	};
	let i = b(e, "drawing");
	if (i) {
		let a = ne(i, t, n);
		return a ? {
			image: a,
			rPr: r
		} : {
			raw: E(e),
			rPr: r
		};
	}
	let a = b(e, "t");
	if (a) return {
		text: a.textContent,
		rPr: r
	};
	let o = e.textContent;
	return o ? {
		text: o,
		rPr: r
	} : {
		raw: E(e),
		rPr: r
	};
}
function E(e) {
	let t = e.attributes.map((e) => ` ${e.name}="${e.value.replace(/"/g, "&quot;")}"`).join(""), n = e.childNodes.map((e) => e.nodeType === 3 ? e.data.replace(/&/g, "&amp;").replace(/</g, "&lt;") : E(e)).join("");
	return `<${e.nodeName}${t}>${n}</${e.nodeName}>`;
}
function ne(t, n, r) {
	let i = t.getElementsByTagNameNS("*", "extent")[0], a = t.getElementsByTagNameNS(e.a, "blip")[0];
	if (!a) return null;
	let o = a.getAttributeNS(g, "embed") || a.getAttribute("r:embed");
	if (!o) return null;
	let s = n.rel(r, o);
	if (!s) return null;
	let c = s.target.startsWith("/") ? s.target.slice(1) : re(r, s.target), l = n.bytes(c);
	if (!l) return null;
	let u = c.split(".").pop() || "png", d = u === "jpg" || u === "jpeg" ? "image/jpeg" : u === "gif" ? "image/gif" : u === "svg" ? "image/svg+xml" : "image/" + u, f = URL.createObjectURL(new Blob([l], { type: d })), p = +(x(i, "cx") || 0), m = +(x(i, "cy") || 0);
	return {
		rId: o,
		part: c,
		url: f,
		w: Math.round(p / _) || 100,
		h: Math.round(m / _) || 100
	};
}
function re(e, t) {
	let n = e.includes("/") ? e.slice(0, e.lastIndexOf("/") + 1) : "", r = [];
	for (let e of (n + t).split("/")) e !== "." && e !== "" && (e === ".." ? r.pop() : r.push(e));
	return r.join("/");
}
function D(e, t, n) {
	let r = c(w(b(e, "pPr")));
	for (let i of e.children) if (i.localName === "r") r.runs.push(T(i, t, n));
	else if (i.localName === "hyperlink") {
		let e = i.getAttributeNS(g, "id") || i.getAttribute("r:id") || "", a = (e ? t.rel(n, e) : void 0)?.target || "";
		for (let o of y(i, "r")) {
			let i = T(o, t, n);
			i.hyperlink = {
				rId: e,
				url: a
			}, r.runs.push(i);
		}
	} else i.localName === "pPr" || i.localName === "bookmarkStart" || i.localName === "bookmarkEnd" || i.localName === "proofErr" || r.runs.push({ raw: E(i) });
	return r;
}
function ie(e, t, n) {
	let r = b(e, "tblGrid"), i = r ? y(r, "gridCol").map((e) => Math.round(+(x(e, "w") || 1e3) / v)) : [], a = [];
	for (let r of y(e, "tr")) {
		let e = [];
		for (let i of y(r, "tc")) {
			let r = b(i, "tcPr"), a = { paras: [] }, o = r ? b(r, "gridSpan") : void 0;
			o && (a.gridSpan = +(x(o, "val") || 1));
			let s = r ? b(r, "vMerge") : void 0;
			s && (a.vMerge = x(s, "val") === "restart" ? "restart" : "continue");
			let l = r ? b(r, "shd") : void 0;
			l && x(l, "fill") && x(l, "fill") !== "auto" && (a.fill = "#" + x(l, "fill"));
			let u = r ? b(r, "tcW") : void 0;
			u && (a.width = Math.round(+(x(u, "w") || 0) / v));
			let d = r ? b(r, "vAlign") : void 0;
			d && (a.vAlign = x(d, "val") === "center" ? "center" : x(d, "val") === "bottom" ? "bottom" : "top"), r && b(r, "tcBorders") && (a.borders = !0);
			let f = r ? b(r, "textDirection") : void 0;
			f && x(f, "val") === "tbRl" && (a.textDirection = "vertical");
			for (let e of y(i, "p")) a.paras.push(D(e, t, n));
			a.paras.length || a.paras.push(c()), e.push(a);
		}
		a.push(e);
	}
	let o = {
		kind: "tbl",
		id: s("t"),
		rows: a,
		widths: i
	}, l = b(e, "tblPr");
	if (l) {
		let t = b(l, "jc");
		t && (o.align = x(t, "val") === "center" ? "center" : x(t, "val") === "right" ? "right" : "left");
		let n = b(l, "tblLayout");
		n && x(n, "type") === "autofit" && (o.autoFit = "content");
		let r = y(e, "tr")[0], i = r ? b(r, "trPr") : void 0;
		i && b(i, "tblHeader") && (o.headerRows = 1);
	}
	return o;
}
function ae(e) {
	let t = [], n = "Calibri", r = 11;
	if (!e.has("word/styles.xml")) return {
		styles: t,
		defaultFont: n,
		defaultSize: r
	};
	let i = e.xml("word/styles.xml"), a = b(i, "docDefaults"), o = C(a ? b(b(a, "rPrDefault") || a, "rPr") : void 0);
	o?.font && (n = o.font), o?.size && (r = o.size);
	for (let e of y(i, "style")) {
		let n = x(e, "styleId") || "", r = {
			id: n,
			name: x(b(e, "name"), "val") || n
		}, i = b(e, "basedOn");
		i && (r.basedOn = x(i, "val") || void 0), r.rPr = C(b(e, "rPr")), r.pPr = w(b(e, "pPr")), t.push(r);
	}
	return {
		styles: t,
		defaultFont: n,
		defaultSize: r
	};
}
function oe(e) {
	let t = {};
	if (!e.has("word/numbering.xml")) return t;
	let n = e.xml("word/numbering.xml"), r = {};
	for (let e of y(n, "abstractNum")) {
		let t = x(e, "abstractNumId") || "", n = y(e, "lvl").find((e) => x(e, "ilvl") === "0") || y(e, "lvl")[0];
		r[t] = n && x(b(n, "numFmt"), "val") || "bullet";
	}
	for (let e of y(n, "num")) {
		let n = +(x(e, "numId") || 0);
		t[n] = r[x(b(e, "abstractNumId"), "val") || ""] || "bullet";
	}
	return t;
}
function se(e) {
	return O(b(e, "sectPr"));
}
function O(e) {
	let t = l();
	if (!e) return t;
	let n = b(e, "pgSz");
	n && (t.pageW = +(x(n, "w") || t.pageW), t.pageH = +(x(n, "h") || t.pageH));
	let r = b(e, "pgMar");
	r && (t.margins = {
		top: +(x(r, "top") || 1440),
		right: +(x(r, "right") || 1440),
		bottom: +(x(r, "bottom") || 1440),
		left: +(x(r, "left") || 1440)
	});
	let i = b(e, "cols");
	return i && (t.columns = Math.max(1, +(x(i, "num") || 1))), t;
}
async function ce(e) {
	let n = await t.open(e), r = n.main || "word/document.xml", i = n.xml(r), a = b(i, "body") || i, o = [], s = se(a);
	for (let e of a.children) if (e.localName === "p") {
		o.push(D(e, n, r));
		let t = b(e, "pPr"), i = t ? b(t, "sectPr") : void 0;
		i && (s = O(i));
	} else e.localName === "tbl" ? o.push(ie(e, n, r)) : e.localName === "sectPr" || o.push({
		...c(),
		raw: E(e)
	});
	o.length || o.push(c());
	let { styles: l, defaultFont: u, defaultSize: d } = ae(n);
	return {
		doc: {
			blocks: o,
			section: s,
			styles: l,
			defaultFont: u,
			defaultSize: d,
			numbering: oe(n),
			format: n.variant("docx")
		},
		pkg: n,
		part: r
	};
}
// src/engine/docx-write.ts
var k = 15, A = 9525, j = (e) => Math.round(e * k), M = "xmlns:w=\"" + e.w + "\" xmlns:r=\"" + e.r + "\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"" + e.a + "\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"";
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
	return e.align && (t += `<w:jc w:val="${e.align === "justify" ? "both" : e.align}"/>`), e.keepNext && (t += "<w:keepNext/>"), e.keepLines && (t += "<w:keepLines/>"), e.outlineLvl != null && (t += `<w:outlineLvl w:val="${e.outlineLvl}"/>`), e.borders && (t += "<w:pBdr><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"1\" w:color=\"auto\"/></w:pBdr>"), e.shading && (t += `<w:shd w:val="clear" w:color="auto" w:fill="${e.shading.replace("#", "").toUpperCase()}"/>`), t ? `<w:pPr>${t}</w:pPr>` : "";
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
	let n = t.widths.map((e) => `<w:gridCol w:w="${j(e)}"/>`).join(""), r = t.rows.some((e) => e.some((e) => e.borders === !1)) ? "" : "<w:tblBorders>" + [
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
			t.gridSpan && t.gridSpan > 1 && (n += `<w:gridSpan w:val="${t.gridSpan}"/>`), t.vMerge && (n += t.vMerge === "restart" ? "<w:vMerge w:val=\"restart\"/>" : "<w:vMerge/>"), t.fill && (n += `<w:shd w:val="clear" w:color="auto" w:fill="${t.fill.replace("#", "").toUpperCase()}"/>`), t.vAlign && (n += `<w:vAlign w:val="${t.vAlign === "center" ? "center" : t.vAlign === "bottom" ? "bottom" : "top"}"/>`), t.textDirection === "vertical" && (n += "<w:textDirection w:val=\"tbRl\"/>");
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
	return e?.underline && a.push("underline"), e?.strike && a.push("line-through"), a.length && n.push("text-decoration-line:" + a.join(" ")), e?.underlineStyle === "double" && n.push("text-decoration-style:double"), e?.underlineStyle === "words" && n.push("text-decoration-style:dotted"), e?.caps === "all" && n.push("text-transform:uppercase"), e?.caps === "small" && n.push("font-variant:small-caps"), e?.hidden && n.push("color:transparent;text-shadow:none"), e?.color && n.push("color:" + e.color), e?.highlight && n.push("background:" + e.highlight), e?.vertAlign === "superscript" && n.push("vertical-align:super;font-size:smaller"), e?.vertAlign === "subscript" && n.push("vertical-align:sub;font-size:smaller"), n.join(";");
}
function be(e, t) {
	let n = [], r = e.pPr?.style ? t.styles.find((t) => t.id === e.pPr.style) : void 0, i = e.pPr?.align || r?.pPr?.align;
	i && n.push("text-align:" + i);
	let a = e.pPr?.spaceBefore ?? r?.pPr?.spaceBefore, o = e.pPr?.spaceAfter ?? r?.pPr?.spaceAfter;
	a != null && n.push("margin-top:" + a + "px"), o != null && n.push("margin-bottom:" + o + "px");
	let s = e.pPr?.indent ?? 0;
	s && n.push("padding-left:" + s + "px"), e.pPr?.line && n.push(e.pPr.line > 4 ? "line-height:" + e.pPr.line + "px" : "line-height:" + e.pPr.line), e.pPr?.pageBreakBefore && n.push("break-before:page;page-break-before:always"), e.pPr?.keepNext && n.push("page-break-after:avoid"), e.pPr?.keepLines && n.push("page-break-inside:avoid"), e.pPr?.borders && n.push("outline:1px solid #888;outline-offset:1px"), e.pPr?.shading && n.push("background:" + e.pPr.shading);
	let c = r?.rPr;
	return c && n.push(B(c, t)), n.join(";");
}
function xe(e, t) {
	if (e.raw != null) return `<span class="raw" contenteditable="false" data-raw="${encodeURIComponent(e.raw)}">[object]</span>`;
	let n = B(e.rPr, t);
	if (e.image) {
		let t = e.image, r = [n];
		t.rotation && r.push(`transform:rotate(${t.rotation}deg)`), t.contrast != null && r.push(`filter:contrast(${t.contrast}%)`), t.brightness != null && r.push(`filter:brightness(${t.brightness}%)`), (t.wrap === "square" || t.wrap === "tight") && r.push("float:left;margin:0 8px 4px 0"), t.wrap === "behind" && r.push("position:absolute;z-index:-1");
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
		let i = e.gridSpan && e.gridSpan > 1 ? ` colspan="${e.gridSpan}"` : "", a = e.vMerge === "continue" ? " data-vmerge=\"continue\"" : e.vMerge === "restart" ? " data-vmerge=\"restart\"" : "", o = e.borders === !1 ? "border:0" : "border:1px solid #888", s = [
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
function H(e) {
	if (!e) return;
	if (e.startsWith("#")) return e.length === 7 ? e : void 0;
	let t = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(e);
	if (!t) return;
	let n = (e) => (+e).toString(16).padStart(2, "0");
	return "#" + n(t[1]) + n(t[2]) + n(t[3]);
}
function Ce(e, t) {
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
	}) : o = Ce(t, i);
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
	} : c();
	r.id = e.getAttribute("data-id") || r.id;
	let i = e.style.textAlign;
	return i && (r.pPr = {
		...r.pPr,
		align: i
	}), r;
}
function we(e, t, n) {
	let r = e.getAttribute("data-id") || void 0;
	if (e.tagName.toLowerCase() === "table") {
		let t = n && n.kind === "tbl" ? n : void 0, i = Array.from(e.querySelectorAll("col")).map((e) => parseInt(e.style.width) || 100), a = Array.from(e.querySelectorAll(":scope > tbody > tr, :scope > tr")).map((e) => Array.from(e.querySelectorAll(":scope > td, :scope > th")).map((e) => {
			let t = e, n = Array.from(t.querySelectorAll(":scope > .p")).map((e) => G(e));
			n.length || n.push(c());
			let r = { paras: n }, i = t.getAttribute("colspan");
			i && +i > 1 && (r.gridSpan = +i);
			let a = t.getAttribute("data-vmerge");
			(a === "restart" || a === "continue") && (r.vMerge = a);
			let o = H(t.style.backgroundColor);
			o && (r.fill = o);
			let s = +(t.getAttribute("data-w") || 0);
			s && (r.width = s);
			let l = t.style.verticalAlign;
			return (l === "middle" || l === "bottom") && (r.vAlign = l === "middle" ? "center" : "bottom"), t.style.writingMode.includes("vertical") && (r.textDirection = "vertical"), (t.style.border === "0px solid rgb(0, 0, 0)" || t.style.border === "0") && (r.borders = !1), r;
		}));
		return {
			kind: "tbl",
			id: r || s("t"),
			rows: a,
			widths: i.length ? i : t?.widths || [100],
			...t?.align ? { align: t.align } : {},
			...t?.autoFit ? { autoFit: t.autoFit } : {},
			...t?.headerRows ? { headerRows: t.headerRows } : {}
		};
	}
	return G(e, n && n.kind === "p" ? n : void 0);
}
// src/engine/layout.ts
function Te(e, t, n, r) {
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
function Ee(e, t) {
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
			n.set(o, e), r.textContent = (a === "lowerLetter" ? String.fromCharCode(96 + e) : a === "upperLetter" ? String.fromCharCode(64 + e) : a === "lowerRoman" ? De(e) : String(e)) + ". ";
		}
	}
}
function De(e) {
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
function Oe(e) {
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
function ke(e, t) {
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
function Ae(e) {
	return {
		source: e,
		text: e.replace(/^\uFEFF/, "").replace(/\r\n?/g, "\n")
	};
}
function je(e) {
	return e.replace(/\r\n?/g, "\n").replace(/\n?$/, "\n");
}
// src/io/html.ts
function Me(e) {
	if (typeof DOMParser > "u") return e.replace(/<[^>]+>/g, "");
	let t = new DOMParser().parseFromString(e, "text/html");
	return t.body?.innerText ?? t.body?.textContent ?? "";
}
function Ne(e, t = "Document") {
	let n = (e) => e.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"), r = e.split(/\r?\n/).map((e) => `<p>${n(e) || "&nbsp;"}</p>`).join("\n");
	return `<!doctype html><html lang="en"><head><meta charset="utf-8"><title>${n(t)}</title></head><body>${r}</body></html>`;
}
// src/services/csv.ts
function K(e) {
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
function Pe(e) {
	return /[",\r\n]/.test(e) ? `"${e.replace(/"/g, "\"\"")}"` : e;
}
function Fe(e, t = ",") {
	return [e.fields, ...e.rows].map((e) => e.map(Pe).join(t)).join("\r\n") + "\r\n";
}
// src/services/autocorrect.ts
var Ie = {
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
function q(e, t = Ie) {
	let n = t[e] ?? t[e.toLowerCase()];
	return n ? /^\p{Lu}/u.test(e) ? n[0].toUpperCase() + n.slice(1) : e === e.toUpperCase() && e.length > 1 ? n.toUpperCase() : n : e;
}
function Le(e, t) {
	return e.replace(/\S+/g, (e) => q(e, t));
}
// src/services/tables.ts
function J(e, t) {
	let n = Math.max(1, Math.floor(e)), r = Math.max(1, Math.floor(t));
	return {
		kind: "tbl",
		id: s("tbl"),
		widths: Array(r).fill(96),
		rows: Array.from({ length: n }, () => Array.from({ length: r }, () => ({ paras: [c()] })))
	};
}
function Re(e, t, n) {
	return e.rows[t]?.[n];
}
function ze(e) {
	let t = e.widths.length || e.rows[0]?.length || 1;
	return e.rows.push(Array.from({ length: t }, () => ({ paras: [c()] }))), e;
}
function Be(e) {
	e.widths.push(96);
	for (let t of e.rows) t.push({ paras: [c()] });
	return e;
}
function Ve(e, t) {
	return e.rows.length > 1 && e.rows.splice(t, 1), e;
}
function He(e, t) {
	if (e.widths.length > 1) {
		e.widths.splice(t, 1);
		for (let n of e.rows) n.splice(t, 1);
	}
	return e;
}
function Ue(e, t, n) {
	let r = J(t, n), i = e.blocks.length && e.blocks[e.blocks.length - 1].kind === "p" ? e.blocks.length - 1 : e.blocks.length;
	return e.blocks.splice(i, 0, r), r;
}
// src/services/fields.ts
function Y(e) {
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
function We(e, t = {}) {
	let n = Y(e), r = t.now ?? /* @__PURE__ */ new Date();
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
function X(e) {
	let t = [], n = /[\p{L}\p{M}\p{N}]+(?:['’\-][\p{L}\p{M}\p{N}]+)*/gu, r;
	for (; r = n.exec(e);) t.push({
		word: r[0],
		start: r.index,
		end: r.index + r[0].length
	});
	return t;
}
function Ge(e) {
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
function Ke(e, t) {
	return X(e).filter((e) => !t.has(e.word.toLocaleLowerCase())).map((e) => ({
		word: e.word,
		start: e.start,
		end: e.end,
		kind: "spelling",
		suggestions: [],
		message: "Not in dictionary"
	}));
}
function Z(e, t) {
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
function qe(e, t, n = 5) {
	return Array.from(t).map((t) => [t, Z(e.toLowerCase(), t.toLowerCase())]).sort((e, t) => e[1] - t[1]).slice(0, n).map((e) => e[0]);
}
// src/services/drawing.ts
function Je(e, t = "") {
	return {
		image: { ...e },
		rPr: {},
		text: t || void 0
	};
}
function Ye(e, t, n = 640) {
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
function Xe(e) {
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
function Ze(e) {
	return e.replace(/\\'[0-9a-f]{2}/gi, (e) => String.fromCharCode(parseInt(e.slice(2), 16))).replace(/\\par[d]?\s?/gi, "\n").replace(/\\tab\s?/gi, "	").replace(/\\u(-?\d+)\??/g, (e, t) => String.fromCharCode(Number(t) < 0 ? Number(t) + 65536 : Number(t))).replace(/\\[a-z]+\d*\s?/gi, "").replace(/[{}]/g, "").replace(/\r\n?/g, "\n").trim();
}
function Qe(e) {
	return `{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 Calibri;}}\\f0\\fs22 ${e.replace(/\\/g, "\\\\").replace(/[{}]/g, (e) => `\\${e}`).replace(/\n/g, "\\par\n").replace(/\t/g, "\\tab ")}}`;
}
// src/services/review.ts
function $e(e, t) {
	return {
		...e,
		raw: e.raw,
		rPr: { ...e.rPr },
		text: t.kind === "delete" ? void 0 : e.text
	};
}
function et(e) {
	return {
		...e,
		accepted: !0
	};
}
function tt(e) {
	return {
		...e,
		accepted: !1
	};
}
function nt(e) {
	let t = [];
	for (let n of e.blocks) if (n.kind === "p") for (let e of n.runs) e.raw?.includes("comment") && t.push({
		id: e.raw,
		author: "",
		date: "",
		text: e.text ?? ""
	});
	return t;
}
function rt(e) {
	return {
		kind: "p",
		id: `review-${Date.now()}`,
		runs: e ? [{ text: e }] : []
	};
}
// src/services/preserve.ts
function it(e, t, n) {
	return e.has(t) || e.set(t, n), e;
}
function at(e) {
	return Array.from(e, (e) => ({
		...e,
		bytes: e.bytes.slice()
	}));
}
// src/services/commands.ts
var ot = class {
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
}, st = class {
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
function Q(e, t, n = 720) {
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
function ct(e, t, n) {
	return Q(e, t, n).length;
}
// src/services/outline.ts
function lt(e) {
	let t = [], n = [];
	for (let r of e.blocks) {
		if (r.kind !== "p") continue;
		let e = $(r);
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
function $(e) {
	if (e.kind !== "p") return 0;
	let t = e.pPr?.style ?? "", n = /^Heading([1-9])$/i.exec(t);
	return n ? Number(n[1]) : 0;
}
// src/services/textOps.ts
function ut(e, t, n, r) {
	return e.slice(0, t) + r + e.slice(n);
}
function dt(e, t) {
	return t === "upper" ? e.toUpperCase() : t === "lower" ? e.toLowerCase() : e.toLowerCase().replace(/(^|\s)(\p{L})/gu, (e, t, n) => t + n.toUpperCase());
}
function ft(e, t) {
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
function pt(e, t) {
	return {
		text: e,
		rPr: { ...t }
	};
}
// src/services/omml.ts
function mt(e) {
	return e.replace(/<m:br\s*\/?>(\s*)/g, "\n").replace(/<[^>]+>/g, "").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&").replace(/\s+/g, " ").trim();
}
function ht(e) {
	return `<m:oMathPara><m:oMath><m:r><m:t>${e.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</m:t></m:r></m:oMath></m:oMathPara>`;
}
// src/services/templates.ts
var gt = class {
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
function _t(e) {
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
function vt(e) {
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
export { ke as $, Ke as A, Ve as B, Ze as C, Xe as D, Ye as E, Be as F, Fe as G, Le as H, ze as I, Ae as J, Me as K, J as L, X as M, We as N, Ge as O, Y as P, Ee as Q, Ue as R, tt as S, Je as T, q as U, Re as V, K as W, Oe as X, je as Y, Te as Z, it as _, mt as a, te as at, nt as b, ft as c, c as ct, lt as d, u as dt, V as et, ct as f, m as ft, at as g, st as h, ht as i, p as it, qe as j, Z as k, dt as l, d as lt, ot as m, ee as mt, vt as n, ge as nt, pt as o, l as ot, Q as p, s as pt, Ne as q, gt as r, ce as rt, ut as s, h as st, _t as t, we as tt, $ as u, f as ut, et as v, Qe as w, rt as x, $e as y, He as z };
