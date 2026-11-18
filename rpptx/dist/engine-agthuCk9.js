import { NS as e, Package as t, esc as n, escText as r, formats as i, newPackage as a } from "@voffice/react-shared";
// src/engine/model.ts
var o = 0, s = (e = "s") => `${e}${++o}_${Math.random().toString(36).slice(2, 7)}`, c = 9525, l = (e) => e / c, u = (e) => Math.round(e * c), d = () => ({
	id: s("sl"),
	shapes: []
}), f = () => ({
	slideW: 12192e3,
	slideH: 6858e3,
	slides: [d()],
	format: "pptx"
});
function p(e, t, n, r, i = "") {
	return {
		id: s("sh"),
		kind: "sp",
		x: e,
		y: t,
		w: n,
		h: r,
		preset: "rect",
		fill: null,
		line: null,
		text: [{
			align: "left",
			runs: i ? [{ text: i }] : []
		}]
	};
}
var m = (e) => JSON.parse(JSON.stringify(e));
e.p, e.a;
var h = e.r, g = (e, t) => e.children.filter((e) => e.localName === t), _ = (e, t) => e.children.find((e) => e.localName === t), v = (e, t) => e ? e.getAttribute(t) : null, y = (e, t) => e.getElementsByTagNameNS("*", t)[0];
function b(e) {
	let t = e.attributes.map((e) => ` ${e.name}="${e.value.replace(/"/g, "&quot;")}"`).join(""), n = e.childNodes.map((e) => e.nodeType === 3 ? e.data.replace(/&/g, "&amp;").replace(/</g, "&lt;") : b(e)).join("");
	return `<${e.nodeName}${t}>${n}</${e.nodeName}>`;
}
function x(e) {
	if (!e) return;
	let t = {}, n = v(e, "sz");
	n && (t.size = n / 100), v(e, "b") === "1" && (t.bold = !0), v(e, "i") === "1" && (t.italic = !0), v(e, "u") && v(e, "u") !== "none" && (t.underline = !0);
	let r = _(e, "solidFill"), i = r && (_(r, "srgbClr") || _(r, "schemeClr"));
	i && _(r, "srgbClr") && (t.color = ("#" + v(i, "val")).toLowerCase());
	let a = _(e, "latin");
	return a && (t.font = v(a, "typeface") || void 0), Object.keys(t).length ? t : void 0;
}
function S(e) {
	if (!e) return [];
	let t = [];
	for (let n of g(e, "p")) {
		let e = { runs: [] }, r = _(n, "pPr"), i = v(r, "algn");
		i && (e.align = i === "ctr" ? "center" : i === "r" ? "right" : "left"), r && (_(r, "buChar") || _(r, "buAutoNum")) && (e.bullet = !0);
		for (let t of g(n, "r")) {
			let n = _(t, "t");
			e.runs.push({
				text: n?.textContent || "",
				rPr: x(_(t, "rPr"))
			});
		}
		let a = _(n, "endParaRPr");
		!e.runs.length && a && e.runs.push({
			text: "",
			rPr: x(a)
		}), t.push(e);
	}
	return t;
}
function C(e) {
	let t = e && _(e, "xfrm"), n = t && _(t, "off"), r = t && _(t, "ext");
	return {
		x: +(v(n, "x") || 0),
		y: +(v(n, "y") || 0),
		w: +(v(r, "cx") || 0),
		h: +(v(r, "cy") || 0),
		rot: t && v(t, "rot") ? v(t, "rot") / 6e4 : void 0
	};
}
function w(e) {
	let t, n;
	if (!e) return {
		fill: t,
		line: n
	};
	_(e, "noFill") && (t = null);
	let r = _(e, "solidFill");
	if (r) {
		let e = _(r, "srgbClr");
		e && (t = { color: ("#" + v(e, "val")).toLowerCase() });
	}
	let i = _(e, "ln");
	if (i) {
		if (_(i, "noFill")) n = null;
		else {
			let e = _(i, "solidFill"), t = e && _(e, "srgbClr");
			n = {
				color: t ? ("#" + v(t, "val")).toLowerCase() : "#000000",
				width: (v(i, "w") || 12700) / 12700
			};
			let r = _(i, "prstDash");
			r && (n.dash = v(r, "val") || void 0);
		}
	}
	return {
		fill: t,
		line: n
	};
}
function T(e, t, n) {
	let r = e.localName, i = _(e, "nvSpPr") || _(e, "nvPicPr") || _(e, "nvCxnSpPr") || _(e, "nvGraphicFramePr"), a = v(i && _(i, "cNvPr"), "name") || "", o = _(e, "spPr"), c = C(o);
	if (r === "pic") {
		let r = y(e, "blip"), i = r && (r.getAttributeNS(h, "embed") || v(r, "r:embed")) || "", o = i ? t.rel(n, i) : void 0, l;
		if (o) {
			let e = o.target.startsWith("/") ? o.target.slice(1) : E(n, o.target), r = t.bytes(e);
			if (r) {
				let t = e.split(".").pop() || "png", n = t === "jpg" || t === "jpeg" ? "image/jpeg" : "image/" + t;
				l = {
					part: e,
					rId: i,
					url: URL.createObjectURL(new Blob([r], { type: n }))
				};
			}
		}
		return {
			id: s("sh"),
			kind: "pic",
			...c,
			image: l,
			name: a
		};
	}
	if (r === "sp" || r === "cxnSp") {
		let t = o && _(o, "prstGeom"), { fill: n, line: i } = w(o);
		return {
			id: s("sh"),
			kind: r === "cxnSp" ? "cxnSp" : "sp",
			...c,
			preset: v(t, "prst") || (r === "cxnSp" ? "line" : "rect"),
			fill: n,
			line: i,
			name: a,
			text: S(_(e, "txBody"))
		};
	}
	return {
		id: s("sh"),
		kind: "graphicFrame",
		...c,
		raw: b(e),
		name: a
	};
}
function E(e, t) {
	let n = e.includes("/") ? e.slice(0, e.lastIndexOf("/") + 1) : "", r = [];
	for (let e of (n + t).split("/")) e !== "." && e !== "" && (e === ".." ? r.pop() : r.push(e));
	return r.join("/");
}
async function D(e) {
	let n = await t.open(e), r = n.xml("ppt/presentation.xml"), i = _(r, "sldSz"), a = {
		slideW: +(v(i, "cx") || 12192e3),
		slideH: +(v(i, "cy") || 6858e3),
		slides: [],
		format: n.variant("pptx")
	}, o = n.relsOf("ppt/presentation.xml"), s = _(r, "sldIdLst"), c = [], l = s ? g(s, "sldId") : [];
	for (let e of l) {
		let t = e.getAttributeNS(h, "id") || v(e, "r:id") || "", r = o.find((e) => e.id === t);
		if (!r) continue;
		let i = r.target.startsWith("/") ? r.target.slice(1) : E("ppt/presentation.xml", r.target);
		c.push(i);
		let s = d();
		if (s.part = i, n.has(i)) {
			let e = _(n.xml(i), "cSld"), t = e && _(e, "bg"), r = t && y(t, "srgbClr");
			r && (s.background = ("#" + v(r, "val")).toLowerCase());
			let a = e && _(e, "spTree");
			if (a) {
				for (let e of a.children) if ([
					"sp",
					"pic",
					"cxnSp",
					"graphicFrame"
				].includes(e.localName)) {
					let t = T(e, n, i);
					t && s.shapes.push(t);
				}
			}
			let o = n.relsOf(i).find((e) => e.type.endsWith("/notesSlide"));
			if (o) {
				let e = E(i, o.target);
				if (n.has(e)) {
					let t = y(n.xml(e), "txBody");
					t && (s.notes = g(t, "p").map((e) => g(e, "r").map((e) => _(e, "t")?.textContent || "").join("")).join("\n"));
				}
			}
		}
		a.slides.push(s);
	}
	return a.slides.length || (a.slides.push(d()), c.push("ppt/slides/slide1.xml")), {
		pres: a,
		pkg: n,
		slideParts: c
	};
}
// src/engine/pptx-write.ts
var O = e.p, k = e.a, A = `xmlns:a="${k}" xmlns:r="${e.r}" xmlns:p="${O}"`, j = (e) => e.replace("#", "").toUpperCase();
function M(e) {
	return e === null ? "<a:noFill/>" : e && e.color ? `<a:solidFill><a:srgbClr val="${j(e.color)}"/></a:solidFill>` : "";
}
function N(e) {
	return e === null ? "<a:ln><a:noFill/></a:ln>" : e && e.color ? `<a:ln w="${Math.round((e.width || 1) * 12700)}"><a:solidFill><a:srgbClr val="${j(e.color)}"/></a:solidFill>${e.dash ? `<a:prstDash val="${e.dash}"/>` : ""}</a:ln>` : "";
}
function P(e) {
	if (!e) return "";
	let t = ["lang=\"en-US\""];
	e.size && t.push(`sz="${Math.round(e.size * 100)}"`), e.bold && t.push("b=\"1\""), e.italic && t.push("i=\"1\""), e.underline && t.push("u=\"sng\"");
	let r = "";
	return e.color && (r += `<a:solidFill><a:srgbClr val="${j(e.color)}"/></a:solidFill>`), e.font && (r += `<a:latin typeface="${n(e.font)}"/>`), `<a:rPr ${t.join(" ")} dirty="0">${r}</a:rPr>`;
}
function F(e) {
	return `<a:p>${e.align || e.bullet ? `<a:pPr${e.align ? ` algn="${e.align === "center" ? "ctr" : e.align === "right" ? "r" : "l"}"` : ""}>${e.bullet ? "<a:buChar char=\"•\"/>" : "<a:buNone/>"}</a:pPr>` : ""}${e.runs.map((e) => `<a:r>${P(e.rPr)}<a:t>${r(e.text)}</a:t></a:r>`).join("") || "<a:endParaRPr lang=\"en-US\"/>"}</a:p>`;
}
function I(e) {
	return `<p:txBody><a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/>${(e && e.length ? e : [{ runs: [] }]).map(F).join("")}</p:txBody>`;
}
function L(e) {
	return `<a:xfrm${e.rot ? ` rot="${Math.round(e.rot * 6e4)}"` : ""}><a:off x="${Math.round(e.x)}" y="${Math.round(e.y)}"/><a:ext cx="${Math.round(e.w)}" cy="${Math.round(e.h)}"/></a:xfrm>`;
}
var R = 1;
function z(e, t, r) {
	let i = ++R;
	if (e.raw != null) return e.raw;
	if (e.kind === "pic" && e.image) {
		let t = e.image.rId;
		return `<p:pic><p:nvPicPr><p:cNvPr id="${i}" name="${n(e.name || "Picture")}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${t}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${L(e)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
	}
	let a = e.kind === "cxnSp" ? "p:cxnSp" : "p:sp", o = e.kind === "cxnSp" ? "p:nvCxnSpPr" : "p:nvSpPr", s = e.kind === "cxnSp" ? "p:cNvCxnSpPr" : "p:cNvSpPr", c = `<a:prstGeom prst="${n(e.preset || "rect")}"><a:avLst/></a:prstGeom>`, l = e.kind === "cxnSp" ? "" : I(e.text);
	return `<${a}><${o}><p:cNvPr id="${i}" name="${n(e.name || "Shape")}"/><${s}/><p:nvPr/></${o}><p:spPr>${L(e)}${c}${M(e.fill)}${N(e.line)}</p:spPr>${l}</${a}>`;
}
function B(e, t, n) {
	R = 1;
	let r = e.background ? `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${j(e.background)}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` : "", i = e.shapes.map((e) => z(e, t, n)).join("");
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sld ${A}><p:cSld>${r}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` + i + "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>";
}
var V = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a:theme xmlns:a="${k}" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`, H = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldMaster ${A}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`, U = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldLayout ${A} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`, W = {
	pres: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
	master: "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
	layout: "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
	slide: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
	theme: "application/vnd.openxmlformats-officedocument.theme+xml"
};
async function G(t, n = {}) {
	let r = i[n.format || t.format || "pptx"] || i.pptx, o = n.pkg || a("ppt/presentation.xml", "", r), s = n.slideParts && n.slideParts.length === t.slides.length ? n.slideParts : t.slides.map((e, t) => `ppt/slides/slide${t + 1}.xml`);
	for (let n = 0; n < t.slides.length; n++) {
		let r = s[n];
		for (let i of t.slides[n].shapes) if (i.kind === "pic" && i.image && i.image.url && (!i.image.part || !o.has(i.image.part))) {
			let t = await fetch(i.image.url), n = new Uint8Array(await t.arrayBuffer()), a = (i.image.part.split(".").pop() || "png").toLowerCase(), s = o.nextMediaName(a === "jpg" ? "jpeg" : a), c = "ppt/media/" + s;
			o.addPart(c, n), i.image.rId = o.addRel(r, e.rel + "/image", "../media/" + s), i.image.part = c;
		}
	}
	if (t.slides.forEach((t, r) => {
		let i = s[r];
		o.addPart(i, B(t, o, i), W.slide), n.pkg || o.addRel(i, e.rel + "/slideLayout", "../slideLayouts/slideLayout1.xml");
	}), !n.pkg) {
		o.addPart("ppt/theme/theme1.xml", V, W.theme), o.addPart("ppt/slideMasters/slideMaster1.xml", H, W.master), o.addPart("ppt/slideLayouts/slideLayout1.xml", U, W.layout), o.addRel("ppt/slideMasters/slideMaster1.xml", e.rel + "/slideLayout", "../slideLayouts/slideLayout1.xml"), o.addRel("ppt/slideMasters/slideMaster1.xml", e.rel + "/theme", "../theme/theme1.xml"), o.addRel("ppt/slideLayouts/slideLayout1.xml", e.rel + "/slideMaster", "../slideMasters/slideMaster1.xml");
		let n = t.slides.map((e, t) => `<p:sldId id="${256 + t}" r:id="rId${t + 2}"/>`).join("");
		o.setPart("ppt/presentation.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:presentation ${A}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${n}</p:sldIdLst><p:sldSz cx="${t.slideW}" cy="${t.slideH}"/><p:notesSz cx="${t.slideH}" cy="${t.slideW}"/></p:presentation>`);
		let r = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"];
		r.push(`<Relationship Id="rId1" Type="${e.rel}/slideMaster" Target="slideMasters/slideMaster1.xml"/>`), t.slides.forEach((t, n) => r.push(`<Relationship Id="rId${n + 2}" Type="${e.rel}/slide" Target="slides/slide${n + 1}.xml"/>`)), r.push("</Relationships>"), o.setPart("ppt/_rels/presentation.xml.rels", r.join("")), o.addRel("", e.rel + "/officeDocument", "ppt/presentation.xml");
	} else if (K(n.pkg) !== t.slides.length) {
		t.slides.map((e, t) => `<p:sldId id="${256 + t}" r:id="rId${t + 2}"/>`).join("");
		let n = o.xml("ppt/presentation.xml").children.find((e) => e.localName === "sldIdLst");
		n && (n.childNodes = [], n.children = []), o.setPart("ppt/presentation.xml", q(o, t));
		let r = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"];
		r.push(`<Relationship Id="rId1" Type="${e.rel}/slideMaster" Target="slideMasters/slideMaster1.xml"/>`), t.slides.forEach((t, n) => r.push(`<Relationship Id="rId${n + 2}" Type="${e.rel}/slide" Target="slides/slide${n + 1}.xml"/>`)), r.push("</Relationships>"), o.setPart("ppt/_rels/presentation.xml.rels", r.join(""));
	}
	return o.save(r.mime);
}
function K(e) {
	if (!e.has("ppt/presentation.xml")) return 0;
	let t = e.xml("ppt/presentation.xml").children.find((e) => e.localName === "sldIdLst");
	return t ? t.children.length : 0;
}
function q(e, t) {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:presentation ${A}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${t.slides.map((e, t) => `<p:sldId id="${256 + t}" r:id="rId${t + 2}"/>`).join("")}</p:sldIdLst><p:sldSz cx="${t.slideW}" cy="${t.slideH}"/><p:notesSz cx="${t.slideH}" cy="${t.slideW}"/></p:presentation>`;
}
// src/engine/shapes.ts
function J(e, t, n, r, i, a) {
	let o = r || "none", s = i || "none", c = i ? a : 0, l = `fill="${o}" stroke="${s}" stroke-width="${c}"`, u = (e) => `<path d="${e}" ${l}/>`, d = (e, t, n, r) => `<rect x="${e}" y="${t}" width="${n}" height="${r}" ${l}/>`, f = (e, t, n, r) => `<ellipse cx="${e}" cy="${t}" rx="${n}" ry="${r}" ${l}/>`, p = (e) => `<polygon points="${e.map((e) => e.join(",")).join(" ")}" ${l}/>`, m = c / 2;
	switch (e) {
		case "ellipse":
		case "oval": return f(t / 2, n / 2, t / 2 - m, n / 2 - m);
		case "triangle": return p([
			[t / 2, m],
			[t - m, n - m],
			[m, n - m]
		]);
		case "rtTriangle": return p([
			[m, m],
			[m, n - m],
			[t - m, n - m]
		]);
		case "diamond": return p([
			[t / 2, m],
			[t - m, n / 2],
			[t / 2, n - m],
			[m, n / 2]
		]);
		case "parallelogram": return p([
			[t * .25, m],
			[t - m, m],
			[t * .75, n - m],
			[m, n - m]
		]);
		case "trapezoid": return p([
			[t * .25, m],
			[t * .75, m],
			[t - m, n - m],
			[m, n - m]
		]);
		case "pentagon": return p([
			[t / 2, m],
			[t - m, n * .4],
			[t * .8, n - m],
			[t * .2, n - m],
			[m, n * .4]
		]);
		case "hexagon": return p([
			[t * .25, m],
			[t * .75, m],
			[t - m, n / 2],
			[t * .75, n - m],
			[t * .25, n - m],
			[m, n / 2]
		]);
		case "octagon": return p([
			[t * .3, m],
			[t * .7, m],
			[t - m, n * .3],
			[t - m, n * .7],
			[t * .7, n - m],
			[t * .3, n - m],
			[m, n * .7],
			[m, n * .3]
		]);
		case "star5": {
			let e = [];
			for (let r = 0; r < 10; r++) {
				let i = r % 2 == 0 ? Math.min(t, n) / 2 - m : Math.min(t, n) / 4.5, a = Math.PI / 5 * r - Math.PI / 2;
				e.push([t / 2 + i * Math.cos(a) * (t / Math.min(t, n)), n / 2 + i * Math.sin(a) * (n / Math.min(t, n))]);
			}
			return p(e);
		}
		case "heart": return u(`M ${t / 2} ${n * .9} C ${t * .1} ${n * .55} ${t * .05} ${n * .2} ${t * .3} ${n * .15} C ${t * .45} ${n * .12} ${t / 2} ${n * .3} ${t / 2} ${n * .35} C ${t / 2} ${n * .3} ${t * .55} ${n * .12} ${t * .7} ${n * .15} C ${t * .95} ${n * .2} ${t * .9} ${n * .55} ${t / 2} ${n * .9} Z`);
		case "rightArrow": return p([
			[m, n * .3],
			[t * .6, n * .3],
			[t * .6, m],
			[t - m, n / 2],
			[t * .6, n - m],
			[t * .6, n * .7],
			[m, n * .7]
		]);
		case "leftArrow": return p([
			[t - m, n * .3],
			[t * .4, n * .3],
			[t * .4, m],
			[m, n / 2],
			[t * .4, n - m],
			[t * .4, n * .7],
			[t - m, n * .7]
		]);
		case "upArrow": return p([
			[t * .3, n - m],
			[t * .3, n * .4],
			[m, n * .4],
			[t / 2, m],
			[t - m, n * .4],
			[t * .7, n * .4],
			[t * .7, n - m]
		]);
		case "downArrow": return p([
			[t * .3, m],
			[t * .3, n * .6],
			[m, n * .6],
			[t / 2, n - m],
			[t - m, n * .6],
			[t * .7, n * .6],
			[t * .7, m]
		]);
		case "chevron": return p([
			[m, m],
			[t * .7, m],
			[t - m, n / 2],
			[t * .7, n - m],
			[m, n - m],
			[t * .3, n / 2]
		]);
		case "homePlate": return p([
			[m, m],
			[t * .7, m],
			[t - m, n / 2],
			[t * .7, n - m],
			[m, n - m]
		]);
		case "roundRect": return `<rect x="${m}" y="${m}" width="${t - c}" height="${n - c}" rx="${Math.min(t, n) * .12}" ${l}/>`;
		case "line":
		case "straightConnector1": return `<line x1="${m}" y1="${n - m}" x2="${t - m}" y2="${m}" stroke="${i || "#000"}" stroke-width="${Math.max(1, c)}"/>`;
		case "bentConnector3": return u(`M ${m} ${n - m} L ${t / 2} ${n - m} L ${t / 2} ${m} L ${t - m} ${m}`);
		case "cloud": return u(`M ${t * .25} ${n * .7} a ${t * .15} ${n * .18} 0 0 1 ${t * .05} ${-n * .35} a ${t * .18} ${n * .2} 0 0 1 ${t * .3} ${-n * .15} a ${t * .18} ${n * .2} 0 0 1 ${t * .3} ${n * .15} a ${t * .15} ${n * .18} 0 0 1 ${t * .05} ${n * .35} Z`);
		default: return d(m, m, t - c, n - c);
	}
}
var Y = [
	"rect",
	"roundRect",
	"ellipse",
	"triangle",
	"rtTriangle",
	"diamond",
	"parallelogram",
	"trapezoid",
	"pentagon",
	"hexagon",
	"octagon",
	"star5",
	"heart",
	"rightArrow",
	"leftArrow",
	"upArrow",
	"downArrow",
	"chevron",
	"homePlate",
	"line",
	"cloud"
], X = (e) => e.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
function Z(e) {
	let t = l(e.x), n = l(e.y), r = l(e.w), i = l(e.h), a = `position:absolute;left:${t}px;top:${n}px;width:${r}px;height:${i}px;${e.rot ? `transform:rotate(${e.rot}deg);` : ""}`;
	if (e.raw != null) return `<div style="${a}border:1px dashed #999;background:#eee"></div>`;
	if (e.kind === "pic" && e.image) return `<div style="${a}"><img src="${e.image.url}" style="width:100%;height:100%;object-fit:fill" draggable="false"></div>`;
	let o = J(e.preset || "rect", r, i, e.fill ? e.fill.color : null, e.line ? e.line.color : null, e.line?.width || 1), s = (e.text || []).map((e) => {
		let t = e.align || "left", n = e.runs.map((e) => {
			let t = [];
			return e.rPr?.bold && t.push("font-weight:bold"), e.rPr?.italic && t.push("font-style:italic"), e.rPr?.underline && t.push("text-decoration:underline"), e.rPr?.size && t.push(`font-size:${e.rPr.size}px`), e.rPr?.color && t.push(`color:${e.rPr.color}`), e.rPr?.font && t.push(`font-family:${e.rPr.font},sans-serif`), `<span style="${t.join(";")}">${X(e.text)}</span>`;
		}).join("");
		return `<div style="text-align:${t};min-height:1em">${e.bullet ? "• " : ""}${n}</div>`;
	}).join(""), c = (e.text || []).some((e) => e.runs.some((e) => e.text));
	return `<div style="${a}"><svg width="${r}" height="${i}" viewBox="0 0 ${r} ${i}" style="position:absolute;inset:0;overflow:visible">${o}</svg>` + (c || e.kind === "sp" ? `<div style="position:absolute;inset:4px;overflow:hidden;color:#000;font-size:18px">${s}</div>` : "") + "</div>";
}
function Q(e, t, n) {
	let r = n / l(t.slideW), i = e.shapes.map(Z).join("");
	return `<div style="position:relative;width:${l(t.slideW)}px;height:${l(t.slideH)}px;background:${e.background || "#fff"};transform:scale(${r});transform-origin:top left">${i}</div>`;
}
export { D as a, u as c, p as d, l as f, G as i, f as l, Y as n, c as o, s as p, J as r, m as s, Q as t, d as u };
