// ../vword/src/engine/model.ts
var e = 0, t = (t = "b") => `${t}${++e}_${Math.random().toString(36).slice(2, 7)}`, n = (e) => ({
	kind: "p",
	id: t("p"),
	runs: [],
	pPr: e
}), r = () => ({
	pageW: 12240,
	pageH: 15840,
	margins: {
		top: 1440,
		right: 1440,
		bottom: 1440,
		left: 1440
	}
}), i = () => ({
	blocks: [n()],
	section: r(),
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
}), a = (e) => e.kind === "p", o = (e) => e.kind === "tbl";
function s(e) {
	return o(e) ? e.rows.map((e) => e.map((e) => e.paras.map(c).join(" ")).join("	")).join("\n") : c(e);
}
function c(e) {
	return e.runs.map((e) => e.text ?? (e.tab ? "	" : e.br ? "\n" : "")).join("");
}
function l(e) {
	return e.blocks.map(s).join("\n");
}
function u(e) {
	let t = l(e).trim();
	return t ? t.split(/\s+/).length : 0;
}
var d = (e) => JSON.parse(JSON.stringify(e)), f = new TextEncoder(), p = new TextDecoder("utf-8"), m = (() => {
	let e = /* @__PURE__ */ new Uint32Array(256);
	for (let t = 0; t < 256; t++) {
		let n = t;
		for (let e = 0; e < 8; e++) n = n & 1 ? 3988292384 ^ n >>> 1 : n >>> 1;
		e[t] = n >>> 0;
	}
	return e;
})();
function h(e) {
	let t = 4294967295;
	for (let n = 0; n < e.length; n++) t = m[(t ^ e[n]) & 255] ^ t >>> 8;
	return (t ^ 4294967295) >>> 0;
}
var g = [
	3,
	4,
	5,
	6,
	7,
	8,
	9,
	10,
	11,
	13,
	15,
	17,
	19,
	23,
	27,
	31,
	35,
	43,
	51,
	59,
	67,
	83,
	99,
	115,
	131,
	163,
	195,
	227,
	258
], _ = [
	0,
	0,
	0,
	0,
	0,
	0,
	0,
	0,
	1,
	1,
	1,
	1,
	2,
	2,
	2,
	2,
	3,
	3,
	3,
	3,
	4,
	4,
	4,
	4,
	5,
	5,
	5,
	5,
	0
], ee = [
	1,
	2,
	3,
	4,
	5,
	7,
	9,
	13,
	17,
	25,
	33,
	49,
	65,
	97,
	129,
	193,
	257,
	385,
	513,
	769,
	1025,
	1537,
	2049,
	3073,
	4097,
	6145,
	8193,
	12289,
	16385,
	24577
], te = [
	0,
	0,
	0,
	0,
	1,
	1,
	2,
	2,
	3,
	3,
	4,
	4,
	5,
	5,
	6,
	6,
	7,
	7,
	8,
	8,
	9,
	9,
	10,
	10,
	11,
	11,
	12,
	12,
	13,
	13
], ne = [
	16,
	17,
	18,
	0,
	8,
	7,
	9,
	6,
	10,
	5,
	11,
	4,
	12,
	3,
	13,
	2,
	14,
	1,
	15
];
function v(e, t) {
	let n = /* @__PURE__ */ new Uint16Array(16), r = /* @__PURE__ */ new Uint16Array(16), i = new Uint16Array(t);
	for (let r = 0; r < t; r++) n[e[r]]++;
	n[0] = 0;
	for (let e = 1; e < 15; e++) r[e + 1] = r[e] + n[e];
	for (let n = 0; n < t; n++) e[n] && (i[r[e[n]]++] = n);
	return {
		counts: n,
		syms: i
	};
}
var y = null, b = null;
function re() {
	if (y) return;
	let e = /* @__PURE__ */ new Uint8Array(288);
	for (let t = 0; t < 144; t++) e[t] = 8;
	for (let t = 144; t < 256; t++) e[t] = 9;
	for (let t = 256; t < 280; t++) e[t] = 7;
	for (let t = 280; t < 288; t++) e[t] = 8;
	y = v(e, 288), b = v((/* @__PURE__ */ new Uint8Array(30)).fill(5), 30);
}
function ie(e, t) {
	let n = 0, r = 0, i = 0, a = new Uint8Array(Math.max(1024, t || e.length * 4)), o = 0, s = (e) => {
		if (o + e <= a.length) return;
		let t = a.length * 2;
		for (; t < o + e;) t *= 2;
		let n = new Uint8Array(t);
		n.set(a.subarray(0, o)), a = n;
	}, c = (t) => {
		for (; i < t;) {
			if (n >= e.length) throw Error("inflate: unexpected end of data");
			r |= e[n++] << i, i += 8;
		}
		let a = r & (1 << t) - 1;
		return r >>>= t, i -= t, a;
	}, l = (e) => {
		let t = 0, n = 0, r = 0;
		for (let i = 1; i < 16; i++) {
			t |= c(1);
			let a = e.counts[i];
			if (t - a < n) return e.syms[r + (t - n)];
			r += a, n += a, n <<= 1, t <<= 1;
		}
		throw Error("inflate: bad code");
	}, u = (e, t) => {
		for (;;) {
			let n = l(e);
			if (n < 256) s(1), a[o++] = n;
			else if (n === 256) return;
			else {
				if (n -= 257, n >= 29) throw Error("inflate: bad length");
				let e = g[n] + c(_[n]), r = l(t), i = ee[r] + c(te[r]);
				if (i > o) throw Error("inflate: distance too far");
				s(e);
				for (let t = 0; t < e; t++) a[o] = a[o - i], o++;
			}
		}
	}, d;
	do {
		d = c(1);
		let t = c(2);
		if (t === 0) {
			r = 0, i = 0;
			let t = e[n] | e[n + 1] << 8;
			n += 4, s(t), a.set(e.subarray(n, n + t), o), o += t, n += t;
		} else if (t === 1) re(), u(y, b);
		else if (t === 2) {
			let e = c(5) + 257, t = c(5) + 1, n = c(4) + 4, r = /* @__PURE__ */ new Uint8Array(19);
			for (let e = 0; e < n; e++) r[ne[e]] = c(3);
			let i = v(r, 19), a = new Uint8Array(e + t), o = 0;
			for (; o < e + t;) {
				let e = l(i);
				if (e < 16) a[o++] = e;
				else {
					let t = 0, n = 0;
					if (e === 16) {
						if (!o) throw Error("inflate: repeat with no prior");
						n = a[o - 1], t = 3 + c(2);
					} else t = e === 17 ? 3 + c(3) : 11 + c(7);
					for (; t--;) a[o++] = n;
				}
			}
			u(v(a.subarray(0, e), e), v(a.subarray(e), t));
		} else throw Error("inflate: bad block type");
	} while (!d);
	return a.slice(0, o);
}
async function x(e, t) {
	let n = new Response(new Blob([e]).stream().pipeThrough(t));
	return new Uint8Array(await n.arrayBuffer());
}
async function ae(e, t) {
	if (typeof DecompressionStream == "function") try {
		return await x(e, new DecompressionStream("deflate-raw"));
	} catch {}
	return ie(e, t);
}
async function oe(e) {
	if (typeof CompressionStream == "function") try {
		return await x(e, new CompressionStream("deflate-raw"));
	} catch {}
	return null;
}
function S(e, t, n, r, i, a, o, s, c) {
	let l = null;
	return {
		name: n,
		method: r,
		csize: i,
		usize: a,
		loff: o,
		crc: c,
		async bytes() {
			if (l) return l;
			let c = s;
			if (c == null) {
				let r = o;
				if (r + 30 > e.length || t.getUint32(r, !0) !== 67324752) throw Error("Damaged entry " + n);
				c = r + 30 + t.getUint16(r + 26, !0) + t.getUint16(r + 28, !0);
			}
			let u = e.subarray(c, Math.min(e.length, c + i)), d;
			if (r === 0) d = u.slice();
			else if (r === 8) d = await ae(u, a);
			else throw Error("Unsupported compression method " + r + " in " + n);
			return l = d, d;
		},
		async text() {
			return p.decode(await this.bytes());
		}
	};
}
var C = (e, t) => t & 2048 ? p.decode(e) : Array.from(e, (e) => String.fromCharCode(e)).join("");
function w(e, t) {
	if (t.name.endsWith("/")) return;
	e.set(t.name, t);
	let n = t.name.replace(/\\/g, "/");
	n !== t.name && !e.has(n) && e.set(n, t);
}
function se(e, t) {
	let n = -1;
	for (let r = e.length - 22; r >= Math.max(0, e.length - 65557); r--) if (t.getUint32(r, !0) === 101010256) {
		n = r;
		break;
	}
	if (n < 0) return null;
	let r = t.getUint16(n + 10, !0), i = t.getUint32(n + 16, !0);
	if (i === 4294967295 || r === 65535) {
		let e = n - 20;
		if (e >= 0 && t.getUint32(e, !0) === 117853008) {
			let n = Number(t.getBigUint64(e + 8, !0));
			r = Number(t.getBigUint64(n + 32, !0)), i = Number(t.getBigUint64(n + 48, !0));
		}
	}
	let a = /* @__PURE__ */ new Map(), o = i;
	for (let n = 0; n < r && !(o + 46 > e.length || t.getUint32(o, !0) !== 33639248); n++) {
		let n = t.getUint16(o + 8, !0), r = t.getUint16(o + 10, !0), i = t.getUint32(o + 16, !0), s = t.getUint32(o + 20, !0), c = t.getUint32(o + 24, !0), l = t.getUint16(o + 28, !0), u = t.getUint16(o + 30, !0), d = t.getUint16(o + 32, !0), f = t.getUint32(o + 42, !0), p = C(e.subarray(o + 46, o + 46 + l), n), m = o + 46 + l, h = Math.min(e.length, m + u);
		for (; m + 4 <= h;) {
			let e = t.getUint16(m, !0), n = t.getUint16(m + 2, !0);
			if (e === 1) {
				let e = m + 4;
				c === 4294967295 && (c = Number(t.getBigUint64(e, !0)), e += 8), s === 4294967295 && (s = Number(t.getBigUint64(e, !0)), e += 8), f === 4294967295 && (f = Number(t.getBigUint64(e, !0)));
			}
			m += 4 + n;
		}
		o += 46 + l + u + d, w(a, S(e, t, p, r, s, c, f, void 0, i));
	}
	return a.size ? a : null;
}
function ce(e, t) {
	let n = /* @__PURE__ */ new Map(), r = 0;
	for (; r + 30 <= e.length;) {
		if (t.getUint32(r, !0) !== 67324752) {
			let t = r + 1;
			for (; t + 4 <= e.length && (e[t] !== 80 || e[t + 1] !== 75 || e[t + 2] !== 3 || e[t + 3] !== 4);) t++;
			if (t + 30 > e.length) break;
			r = t;
			continue;
		}
		let i = t.getUint16(r + 6, !0), a = t.getUint16(r + 8, !0), o = t.getUint32(r + 18, !0), s = t.getUint32(r + 22, !0), c = t.getUint16(r + 26, !0), l = t.getUint16(r + 28, !0), u = C(e.subarray(r + 30, r + 30 + c), i), d = r + 30 + c + l, f = d + o;
		if (i & 8 || o === 4294967295 || o === 0 && a === 8) {
			let n = d, r = !1;
			for (; n + 16 <= e.length;) {
				if (e[n] === 80 && e[n + 1] === 75) {
					let e = t.getUint32(n, !0);
					if (e === 134695760 && t.getUint32(n + 8, !0) === n - d) {
						o = n - d, s = t.getUint32(n + 12, !0), f = n + 16, r = !0;
						break;
					}
					if (e === 67324752 || e === 33639248) {
						o = n - d, s = 0, f = n, r = !0;
						break;
					}
				}
				n++;
			}
			r || (o = e.length - d, s = 0, f = e.length);
		}
		if (d > e.length) break;
		w(n, S(e, t, u, a, o, s, r, d)), r = Math.max(f, r + 30);
	}
	return n;
}
async function le(e) {
	let t = e instanceof Uint8Array ? e : new Uint8Array(e), n = new DataView(t.buffer, t.byteOffset, t.byteLength);
	if (t.length < 22) throw Error("This file is not a valid Office Open XML (ZIP) package.");
	if (n.getUint32(0, !0) !== 67324752 && n.getUint32(0, !0) === 3759263696) {
		let e = new TextDecoder("utf-16le").decode(t.subarray(0, Math.min(t.length, 1 << 20))), n = /EncryptedPackage/.test(e), r = /* @__PURE__ */ Error(n ? "This file is password-protected." : "This is a binary Office 97–2003 file, not an Office Open XML package.");
		throw r.code = "ole", r.encrypted = n, r.bytes = t, r;
	}
	let r = null;
	try {
		r = se(t, n);
	} catch {
		r = null;
	}
	if (!r) {
		try {
			r = ce(t, n);
		} catch {
			r = null;
		}
		if (!r || !r.size) throw Error(n.getUint32(0, !0) === 67324752 ? "This file is damaged and could not be repaired." : "This file is not a valid Office Open XML (ZIP) package.");
		r.repaired = !0;
	}
	return r;
}
function ue(e) {
	return {
		time: e.getHours() << 11 | e.getMinutes() << 5 | Math.floor(e.getSeconds() / 2),
		date: e.getFullYear() - 1980 << 9 | e.getMonth() + 1 << 5 | e.getDate()
	};
}
async function de(e, t) {
	let n = [], r = [], i = 0, { time: a, date: o } = ue(/* @__PURE__ */ new Date());
	for (let t of e) {
		let e = f.encode(t.name), s = typeof t.data == "string" ? f.encode(t.data) : t.data, c = h(s), l = null;
		!t.store && s.length > 64 && (l = await oe(s));
		let u = l && l.length < s.length ? 8 : 0, d = u === 8 ? l : s, p = new Uint8Array(30 + e.length), m = new DataView(p.buffer);
		m.setUint32(0, 67324752, !0), m.setUint16(4, 20, !0), m.setUint16(6, 2048, !0), m.setUint16(8, u, !0), m.setUint16(10, a, !0), m.setUint16(12, o, !0), m.setUint32(14, c, !0), m.setUint32(18, d.length, !0), m.setUint32(22, s.length, !0), m.setUint16(26, e.length, !0), m.setUint16(28, 0, !0), p.set(e, 30), n.push(p, d);
		let g = new Uint8Array(46 + e.length), _ = new DataView(g.buffer);
		_.setUint32(0, 33639248, !0), _.setUint16(4, 20, !0), _.setUint16(6, 20, !0), _.setUint16(8, 2048, !0), _.setUint16(10, u, !0), _.setUint16(12, a, !0), _.setUint16(14, o, !0), _.setUint32(16, c, !0), _.setUint32(20, d.length, !0), _.setUint32(24, s.length, !0), _.setUint16(28, e.length, !0), _.setUint32(42, i, !0), g.set(e, 46), r.push(g), i += p.length + d.length;
	}
	let s = r.reduce((e, t) => e + t.length, 0), c = /* @__PURE__ */ new Uint8Array(22), l = new DataView(c.buffer);
	return l.setUint32(0, 101010256, !0), l.setUint16(8, e.length, !0), l.setUint16(10, e.length, !0), l.setUint32(12, s, !0), l.setUint32(16, i, !0), new Blob([
		...n,
		...r,
		c
	], { type: t || "application/zip" });
}
// ../voffice-shared/src/core/xml.ts
var fe = {
	lt: "<",
	gt: ">",
	amp: "&",
	quot: "\"",
	apos: "'"
};
function pe(e) {
	return e.indexOf("&") < 0 ? e : e.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (e, t) => {
		if (t[0] === "#") {
			let n = t[1] === "x" || t[1] === "X" ? parseInt(t.slice(2), 16) : parseInt(t.slice(1), 10);
			if (!(n >= 0 && n <= 1114111)) return e;
			try {
				return String.fromCodePoint(n);
			} catch {
				return e;
			}
		}
		return fe[t] == null ? e : fe[t];
	});
}
var T = (e) => (e = String(e).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"), /[\t\n\r]/.test(e) ? e.replace(/\t/g, "&#9;").replace(/\n/g, "&#10;").replace(/\r/g, "&#13;") : e), me = /* @__PURE__ */ RegExp("[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F\\uFFFE\\uFFFF]", "g"), he = /* @__PURE__ */ RegExp("[\\x00-\\x08\\x0B-\\x1F\\uFFFE\\uFFFF]", "g"), ge = (e, t) => String(e).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/_x([0-9A-Fa-f]{4})_/g, "_x005F_x$1_").replace(t ? me : he, (e) => "_x" + e.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0") + "_"), E = class {
	data;
	parentNode;
	constructor(e, t) {
		this.data = e, this.parentNode = t;
	}
	get nodeType() {
		return 3;
	}
	get textContent() {
		return this.data;
	}
	get nodeValue() {
		return this.data;
	}
	get localName() {
		return null;
	}
	remove() {
		let e = this.parentNode;
		if (!e) return;
		let t = e.childNodes.indexOf(this);
		t >= 0 && e.childNodes.splice(t, 1), this.parentNode = null;
	}
}, D = class e {
	nodeName;
	prefix;
	localName;
	attrs = null;
	children = [];
	childNodes = [];
	parentNode;
	constructor(e, t) {
		this.nodeName = e;
		let n = e.indexOf(":");
		this.prefix = n < 0 ? null : e.slice(0, n), this.localName = n < 0 ? e : e.slice(n + 1), this.parentNode = t;
	}
	get nodeType() {
		return 1;
	}
	get tagName() {
		return this.nodeName;
	}
	getAttribute(e) {
		let t = this.attrs;
		if (!t) return null;
		let n = t[e];
		return n === void 0 ? null : n;
	}
	hasAttribute(e) {
		return this.getAttribute(e) != null;
	}
	setAttribute(e, t) {
		let n = this.attrs ??= Object.create(null);
		n[e] = String(t);
	}
	removeAttribute(e) {
		this.attrs && delete this.attrs[e];
	}
	get attributes() {
		let e = [];
		if (this.attrs) for (let t in this.attrs) {
			let n = t.indexOf(":");
			e.push({
				name: t,
				prefix: n < 0 ? null : t.slice(0, n),
				localName: n < 0 ? t : t.slice(n + 1),
				value: this.attrs[t]
			});
		}
		return e;
	}
	lookupNamespaceURI(e) {
		let t = e ? "xmlns:" + e : "xmlns";
		for (let e = this; e && e.nodeType === 1; e = e.parentNode) {
			let n = e.attrs;
			if (n && n[t] !== void 0) return n[t];
		}
		return e === "xml" ? "http://www.w3.org/XML/1998/namespace" : null;
	}
	get namespaceURI() {
		return this.lookupNamespaceURI(this.prefix);
	}
	getAttributeNS(e, t) {
		let n = this.attrs;
		if (!n) return null;
		if (!e) {
			let e = n[t];
			return e === void 0 ? null : e;
		}
		for (let r in n) {
			let i = r.indexOf(":");
			if (i < 0 || r.slice(i + 1) !== t) continue;
			let a = r.slice(0, i);
			if (a !== "xmlns" && this.lookupNamespaceURI(a) === e) return n[r];
		}
		return null;
	}
	get textContent() {
		if (this.childNodes.length === 1 && this.childNodes[0].nodeType === 3) return this.childNodes[0].data;
		let e = "", t = (n) => {
			for (let r of n.childNodes) e += r.nodeType === 3 ? r.data : (t(r), "");
		};
		return t(this), e;
	}
	set textContent(e) {
		this.children = [], this.childNodes = [new E(String(e), this)];
	}
	get firstElementChild() {
		return this.children[0] || null;
	}
	get firstChild() {
		return this.childNodes[0] || null;
	}
	get parentElement() {
		return this.parentNode && this.parentNode.nodeType === 1 ? this.parentNode : null;
	}
	getElementsByTagName(e) {
		let t = [], n = e === "*", r = (i) => {
			for (let a of i.children) (n || a.nodeName === e) && t.push(a), a.children.length && r(a);
		};
		return r(this), t;
	}
	getElementsByTagNameNS(e, t) {
		let n = [], r = (i) => {
			for (let a of i.children) t !== "*" && a.localName !== t || e !== "*" && a.namespaceURI !== e || n.push(a), a.children.length && r(a);
		};
		return r(this), n;
	}
	appendChild(e) {
		return e.parentNode && e.remove(), e.parentNode = this, this.childNodes.push(e), e.nodeType === 1 && this.children.push(e), e;
	}
	remove() {
		let e = this.parentNode;
		if (!e) return;
		let t = e.childNodes.indexOf(this);
		t >= 0 && e.childNodes.splice(t, 1), t = e.children.indexOf(this), t >= 0 && e.children.splice(t, 1), this.parentNode = null;
	}
	replaceWith(...e) {
		let t = this.parentNode;
		if (!t) return;
		for (let t of e) t.parentNode && t.remove();
		let n = t.childNodes.indexOf(this);
		t.childNodes.splice(n, 1, ...e);
		for (let n of e) n.parentNode = t;
		t.children = t.childNodes.filter((e) => e.nodeType === 1), this.parentNode = null;
	}
	cloneNode(t) {
		let n = new e(this.nodeName, null);
		if (this.attrs && (n.attrs = { ...this.attrs }), t) for (let e of this.childNodes) n.appendChild(e.nodeType === 3 ? new E(e.data, n) : e.cloneNode(!0));
		return n;
	}
}, O = /([^\s=\/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
function _e(e, t) {
	O.lastIndex = 0;
	let n;
	for (; n = O.exec(e);) {
		let e = t.attrs ??= Object.create(null), r = n[2] === void 0 ? n[3] : n[2];
		e[n[1]] = r.indexOf("&") < 0 ? r : pe(r);
	}
}
var ve = /* @__PURE__ */ new Set([
	"t",
	"v",
	"f",
	"formula",
	"formula1",
	"formula2",
	"oddHeader",
	"oddFooter",
	"evenHeader",
	"evenFooter",
	"firstHeader",
	"firstFooter",
	"text",
	"Data",
	"Font",
	"B",
	"I",
	"U",
	"S",
	"Sup",
	"Sub"
]);
function k(e, t) {
	e && e.charCodeAt(0) === 65279 && (e = e.slice(1));
	let n = !!t?.keepWhitespace, r = !!t?.strict, i = new D("#document", null), a = i, o = 0, s = e.length, c = (e) => {
		a !== i && (n || !/^\s*$/.test(e) || ve.has(a.localName)) && a.childNodes.push(new E(pe(e), a));
	};
	for (; o < s;) {
		let t = e.indexOf("<", o);
		if (t < 0) {
			c(e.slice(o));
			break;
		}
		t > o && c(e.slice(o, t));
		let n = e.charCodeAt(t + 1);
		if (n === 47) {
			let n = e.indexOf(">", t);
			if (n < 0) throw Error("XML parse error: unterminated end tag");
			let s = e.slice(t + 2, n).trim(), c = a;
			if (r && (a === i || a.nodeName !== s)) throw Error("XML parse error: mismatched end tag " + s);
			for (; c && c !== i && c.nodeName !== s;) c = c.parentNode;
			c && c !== i && (a = c.parentNode), o = n + 1;
		} else if (n === 33) {
			if (e.startsWith("<!--", t)) {
				let n = e.indexOf("-->", t + 4);
				o = n < 0 ? s : n + 3;
			} else if (e.startsWith("<![CDATA[", t)) {
				let n = e.indexOf("]]>", t + 9), r = e.slice(t + 9, n < 0 ? s : n);
				a !== i && a.childNodes.push(new E(r, a)), o = n < 0 ? s : n + 3;
			} else {
				let n = 0, r = t + 2;
				for (; r < s; r++) {
					let t = e[r];
					if (t === "[") n++;
					else if (t === "]") n--;
					else if (t === ">" && n <= 0) break;
				}
				o = r + 1;
			}
		} else if (n === 63) {
			let n = e.indexOf("?>", t + 2);
			o = n < 0 ? s : n + 2;
		} else {
			let n = t + 1, r = 0;
			for (; n < s; n++) {
				let t = e.charCodeAt(n);
				if (r) t === r && (r = 0);
				else if (t === 34 || t === 39) r = t;
				else if (t === 62) break;
			}
			if (n >= s) throw Error("XML parse error: unterminated start tag");
			let i = t + 1;
			for (; i < n;) {
				let t = e.charCodeAt(i);
				if (t === 32 || t === 9 || t === 10 || t === 13 || t === 47 || t === 62) break;
				i++;
			}
			let c = e.slice(t + 1, i), l = e.charCodeAt(n - 1) === 47, u = new D(c, a), d = l ? n - 1 : n;
			i < d && _e(e.slice(i, d), u), a.childNodes.push(u), a.children.push(u), l || (a = u), o = n + 1;
		}
	}
	let l = i.children[0];
	if (!l) throw Error("XML parse error: no root element");
	if (r && (a !== i || i.children.length !== 1)) throw Error("XML parse error: incomplete document");
	return l;
}
function A(e) {
	return e.length >= 2 && e[0] === 255 && e[1] === 254 ? new TextDecoder("utf-16le").decode(e.subarray(2)) : e.length >= 2 && e[0] === 254 && e[1] === 255 ? new TextDecoder("utf-16be").decode(e.subarray(2)) : e.length >= 4 && e[0] === 0 && e[1] === 60 && e[2] === 0 && e[3] === 63 ? new TextDecoder("utf-16be").decode(e) : e.length >= 4 && e[0] === 60 && e[1] === 0 && e[2] === 63 && e[3] === 0 ? new TextDecoder("utf-16le").decode(e) : new TextDecoder("utf-8").decode(e);
}
// ../voffice-shared/src/core/opc.ts
var j = Object.freeze({
	rel: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	strictRel: "http://purl.oclc.org/ooxml/officeDocument/relationships",
	pkg: "http://schemas.openxmlformats.org/package/2006/relationships",
	ct: "http://schemas.openxmlformats.org/package/2006/content-types",
	mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
	w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
	p: "http://schemas.openxmlformats.org/presentationml/2006/main",
	a: "http://schemas.openxmlformats.org/drawingml/2006/main",
	s: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
	r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
}), M = Object.freeze(Object.fromEntries([
	[
		"docx",
		"wordprocessingml.document",
		"ms-word.document",
		!1,
		!1
	],
	[
		"docm",
		"wordprocessingml.document",
		"ms-word.document",
		!0,
		!1
	],
	[
		"dotx",
		"wordprocessingml.template",
		"ms-word.template",
		!1,
		!0
	],
	[
		"dotm",
		"wordprocessingml.template",
		"ms-word.template",
		!0,
		!0
	],
	[
		"xlsx",
		"spreadsheetml.sheet",
		"ms-excel.sheet",
		!1,
		!1
	],
	[
		"xlsm",
		"spreadsheetml.sheet",
		"ms-excel.sheet",
		!0,
		!1
	],
	[
		"xltx",
		"spreadsheetml.template",
		"ms-excel.template",
		!1,
		!0
	],
	[
		"xltm",
		"spreadsheetml.template",
		"ms-excel.template",
		!0,
		!0
	],
	[
		"pptx",
		"presentationml.presentation",
		"ms-powerpoint.presentation",
		!1,
		!1
	],
	[
		"pptm",
		"presentationml.presentation",
		"ms-powerpoint.presentation",
		!0,
		!1
	],
	[
		"ppsx",
		"presentationml.slideshow",
		"ms-powerpoint.slideshow",
		!1,
		!1
	],
	[
		"ppsm",
		"presentationml.slideshow",
		"ms-powerpoint.slideshow",
		!0,
		!1
	],
	[
		"potx",
		"presentationml.template",
		"ms-powerpoint.template",
		!1,
		!0
	],
	[
		"potm",
		"presentationml.template",
		"ms-powerpoint.template",
		!0,
		!0
	]
].map(([e, t, n, r, i]) => {
	let a = r ? "application/vnd." + n : "application/vnd.openxmlformats-officedocument." + t, o = e, s = r;
	return [o, Object.freeze({
		ext: o,
		macro: s,
		template: i,
		family: t.split(".")[0],
		contentType: a + (s ? o === "dotm" ? ".macroEnabledTemplate.main+xml" : ".macroEnabled.main+xml" : ".main+xml"),
		mime: a + (s ? ".macroEnabled.12" : "")
	})];
}))), N = "[Content_Types].xml", P = (e) => {
	let t = e.lastIndexOf("/");
	return (t < 0 ? "" : e.slice(0, t + 1)) + "_rels/" + (t < 0 ? e : e.slice(t + 1)) + ".rels";
};
function F(e, t) {
	if (t.startsWith("/")) return t.slice(1);
	let n = ((e.includes("/") ? e.slice(0, e.lastIndexOf("/") + 1) : "") + t).split("/"), r = [];
	for (let e of n) e !== "." && e !== "" && (e === ".." ? r.pop() : r.push(e));
	return r.join("/");
}
var I = (e) => e.replace(j.strictRel + "/", j.rel + "/"), L = class e {
	parts = /* @__PURE__ */ new Map();
	edited = /* @__PURE__ */ new Set();
	textCache = /* @__PURE__ */ new Map();
	xmlCache = /* @__PURE__ */ new Map();
	relsCache = /* @__PURE__ */ new Map();
	defaults = /* @__PURE__ */ new Map();
	overrides = /* @__PURE__ */ new Map();
	ctDirty = !1;
	relsDirty = /* @__PURE__ */ new Set();
	main = "";
	repaired = !1;
	static async open(t) {
		let n = await le(t), r = new e();
		r.repaired = !!n.repaired;
		for (let [e, t] of n) r.parts.set(e, await t.bytes());
		let i = r.parts.get(N);
		if (!i) throw Error("This package has no [Content_Types].xml part.");
		let a = k(A(i));
		for (let e of a.children) e.localName === "Default" ? r.defaults.set(e.getAttribute("Extension") || "", e.getAttribute("ContentType") || "") : e.localName === "Override" && r.overrides.set((e.getAttribute("PartName") || "").replace(/^\//, ""), e.getAttribute("ContentType") || "");
		let o = r.relsOf("").find((e) => I(e.type) === j.rel + "/officeDocument");
		return o && (r.main = F("", o.target)), r;
	}
	has(e) {
		return this.parts.has(e);
	}
	bytes(e) {
		return this.parts.get(e);
	}
	text(e) {
		let t = this.textCache.get(e);
		if (t === void 0) {
			let n = this.parts.get(e);
			if (!n) throw Error("No such part: " + e);
			t = A(n), this.textCache.set(e, t);
		}
		return t;
	}
	xml(e) {
		let t = this.xmlCache.get(e);
		return t || (t = k(this.text(e)), this.xmlCache.set(e, t)), t;
	}
	type(e) {
		let t = this.overrides.get(e);
		if (t) return t;
		let n = e.lastIndexOf(".");
		return (n < 0 ? "" : this.defaults.get(e.slice(n + 1))) || "";
	}
	relsOf(e) {
		let t = e || "", n = this.relsCache.get(t);
		if (n) return n;
		n = [];
		let r = this.parts.get(P(e));
		if (r) {
			let e = k(A(r));
			for (let t of e.children) t.localName === "Relationship" && n.push({
				id: t.getAttribute("Id") || "",
				type: I(t.getAttribute("Type") || ""),
				target: t.getAttribute("Target") || "",
				mode: t.getAttribute("TargetMode") || void 0
			});
		}
		return this.relsCache.set(t, n), n;
	}
	rel(e, t) {
		return this.relsOf(e).find((e) => e.id === t);
	}
	related(e, t) {
		return this.relsOf(e).filter((e) => e.type === t && e.mode !== "External").map((t) => F(e, t.target));
	}
	setPart(e, t) {
		this.parts.set(e, typeof t == "string" ? new TextEncoder().encode(t) : t), this.edited.add(e), this.textCache.delete(e), this.xmlCache.delete(e);
	}
	addPart(e, t, n) {
		if (this.setPart(e, t), n) this.overrides.set(e, n), this.ctDirty = !0;
		else {
			let t = e.includes(".") ? e.slice(e.lastIndexOf(".") + 1) : "";
			t && !this.defaults.has(t) && (this.defaults.set(t, ye(t)), this.ctDirty = !0);
		}
	}
	removePart(e) {
		this.parts.delete(e), this.edited.delete(e), this.textCache.delete(e), this.xmlCache.delete(e), this.overrides.delete(e) && (this.ctDirty = !0);
		let t = P(e);
		this.parts.delete(t) && this.relsDirty.delete(t);
	}
	addRel(e, t, n, r) {
		let i = this.relsOf(e).slice(), a = 0;
		for (let e of i) {
			let t = /^rId(\d+)$/.exec(e.id);
			t && (a = Math.max(a, +t[1]));
		}
		let o = "rId" + (a + 1);
		i.push({
			id: o,
			type: t,
			target: n,
			mode: r
		});
		let s = e || "";
		return this.relsCache.set(s, i), this.relsDirty.add(s), o;
	}
	nextMediaName(e, t = "image") {
		let n = 1;
		for (; this.parts.has(`word/media/${t}${n}.${e}`) || this.parts.has(`ppt/media/${t}${n}.${e}`) || this.parts.has(`xl/media/${t}${n}.${e}`);) n++;
		return `${t}${n}.${e}`;
	}
	variant(e) {
		let t = this.type(this.main);
		return Object.values(M).find((e) => e.contentType === t)?.ext || e;
	}
	writeContentTypes() {
		if (!this.ctDirty) return;
		let e = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", "<Types xmlns=\"" + j.ct + "\">"];
		for (let [t, n] of [...this.defaults].sort()) e.push(`<Default Extension="${T(t)}" ContentType="${T(n)}"/>`);
		for (let [t, n] of [...this.overrides].sort()) e.push(`<Override PartName="/${T(t)}" ContentType="${T(n)}"/>`);
		e.push("</Types>"), this.parts.set(N, new TextEncoder().encode(e.join("")));
	}
	writeRels() {
		for (let e of this.relsDirty) {
			let t = this.relsCache.get(e) || [], n = e, r = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", "<Relationships xmlns=\"" + j.pkg + "\">"];
			for (let e of t) r.push(`<Relationship Id="${T(e.id)}" Type="${T(e.type)}" Target="${T(e.target)}"${e.mode ? ` TargetMode="${T(e.mode)}"` : ""}/>`);
			r.push("</Relationships>");
			let i = P(n);
			this.parts.set(i, new TextEncoder().encode(r.join(""))), this.overrides.has(i) || (this.defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml"), this.ctDirty = !0);
		}
	}
	async save(e) {
		return this.writeRels(), this.writeContentTypes(), de([...this.parts].map(([e, t]) => ({
			name: e,
			data: t
		})), e || "application/zip");
	}
};
function ye(e) {
	return {
		xml: "application/xml",
		rels: "application/vnd.openxmlformats-package.relationships+xml",
		png: "image/png",
		jpeg: "image/jpeg",
		jpg: "image/jpeg",
		gif: "image/gif",
		bmp: "image/bmp",
		tiff: "image/tiff",
		svg: "image/svg+xml",
		emf: "image/x-emf",
		wmf: "image/x-wmf",
		bin: "application/vnd.openxmlformats-officedocument.oleObject"
	}[e.toLowerCase()] || "application/octet-stream";
}
function be(e, t, n, r = {}) {
	let i = new L();
	i.main = e, i.defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml"), i.defaults.set("xml", "application/xml");
	for (let [e, t] of Object.entries(r)) i.defaults.set(e, t);
	return i.overrides.set(e, n.contentType), i.setPart(e, t), i.addRel("", j.rel + "/officeDocument", e), i;
}
// ../vword/src/engine/docx-read.ts
j.w;
var R = j.r, z = 9525, B = 15, V = (e, t) => e.children.filter((e) => e.localName === t), H = (e, t) => e.children.find((e) => e.localName === t), U = (e, t) => {
	if (!e) return null;
	let n = e.getAttribute(t);
	if (n != null) return n;
	let r = e.attributes.find((e) => e.localName === t);
	return r ? r.value : null;
}, W = (e) => e != null && e !== "0" && e !== "false" && e !== "none";
function G(e) {
	if (!e) return;
	let t = {};
	H(e, "b") && (t.bold = W(U(H(e, "b"), "val") || "1")), H(e, "i") && (t.italic = W(U(H(e, "i"), "val") || "1")), H(e, "u") && (t.underline = !0), H(e, "strike") && (t.strike = W(U(H(e, "strike"), "val") || "1"));
	let n = H(e, "rFonts");
	n && (t.font = U(n, "ascii") || U(n, "hAnsi") || void 0);
	let r = H(e, "sz");
	r && (t.size = (U(r, "val") || 22) / 2);
	let i = H(e, "color");
	i && U(i, "val") && U(i, "val") !== "auto" && (t.color = ("#" + U(i, "val")).toLowerCase());
	let a = H(e, "highlight");
	a && (t.highlight = U(a, "val") || void 0);
	let o = H(e, "vertAlign");
	return o && (t.vertAlign = U(o, "val") || void 0), Object.keys(t).length ? t : void 0;
}
function xe(e) {
	if (!e) return;
	let t = {};
	H(e, "pageBreakBefore") && (t.pageBreakBefore = !0);
	let n = H(e, "jc");
	if (n) {
		let e = U(n, "val") || "left";
		t.align = e === "both" ? "justify" : e === "start" ? "left" : e === "end" ? "right" : e;
	}
	let r = H(e, "pStyle");
	r && (t.style = U(r, "val") || void 0);
	let i = H(e, "numPr");
	if (i) {
		let e = H(i, "ilvl"), n = H(i, "numId");
		n && (t.numId = +(U(n, "val") || 0)), e && (t.ilvl = +(U(e, "val") || 0));
	}
	let a = H(e, "ind");
	if (a) {
		let e = U(a, "left") || U(a, "start");
		e && (t.indent = Math.round(+e / B));
		let n = U(a, "hanging");
		n && (t.hanging = Math.round(+n / B));
	}
	let o = H(e, "spacing");
	if (o) {
		let e = U(o, "before"), n = U(o, "after"), r = U(o, "line");
		e && (t.spaceBefore = Math.round(+e / B)), n && (t.spaceAfter = Math.round(+n / B)), r && U(o, "lineRule") !== "auto" ? t.line = Math.round(+r / B) : r && (t.line = r / 240);
	}
	return Object.keys(t).length ? t : void 0;
}
function Se(e, t, n) {
	let r = G(H(e, "rPr"));
	if (H(e, "br")) return {
		br: !0,
		rPr: r
	};
	if (H(e, "tab")) return {
		tab: !0,
		rPr: r
	};
	let i = H(e, "drawing");
	if (i) {
		let a = Ce(i, t, n);
		return a ? {
			image: a,
			rPr: r
		} : {
			raw: K(e),
			rPr: r
		};
	}
	let a = H(e, "t");
	if (a) return {
		text: a.textContent,
		rPr: r
	};
	let o = e.textContent;
	return o ? {
		text: o,
		rPr: r
	} : {
		raw: K(e),
		rPr: r
	};
}
function K(e) {
	let t = e.attributes.map((e) => ` ${e.name}="${e.value.replace(/"/g, "&quot;")}"`).join(""), n = e.childNodes.map((e) => e.nodeType === 3 ? e.data.replace(/&/g, "&amp;").replace(/</g, "&lt;") : K(e)).join("");
	return `<${e.nodeName}${t}>${n}</${e.nodeName}>`;
}
function Ce(e, t, n) {
	let r = e.getElementsByTagNameNS("*", "extent")[0], i = e.getElementsByTagNameNS(j.a, "blip")[0];
	if (!i) return null;
	let a = i.getAttributeNS(R, "embed") || i.getAttribute("r:embed");
	if (!a) return null;
	let o = t.rel(n, a);
	if (!o) return null;
	let s = o.target.startsWith("/") ? o.target.slice(1) : we(n, o.target), c = t.bytes(s);
	if (!c) return null;
	let l = s.split(".").pop() || "png", u = l === "jpg" || l === "jpeg" ? "image/jpeg" : l === "gif" ? "image/gif" : l === "svg" ? "image/svg+xml" : "image/" + l, d = URL.createObjectURL(new Blob([c], { type: u })), f = +(U(r, "cx") || 0), p = +(U(r, "cy") || 0);
	return {
		rId: a,
		part: s,
		url: d,
		w: Math.round(f / z) || 100,
		h: Math.round(p / z) || 100
	};
}
function we(e, t) {
	let n = e.includes("/") ? e.slice(0, e.lastIndexOf("/") + 1) : "", r = [];
	for (let e of (n + t).split("/")) e !== "." && e !== "" && (e === ".." ? r.pop() : r.push(e));
	return r.join("/");
}
function Te(e, t, r) {
	let i = n(xe(H(e, "pPr")));
	for (let n of e.children) if (n.localName === "r") i.runs.push(Se(n, t, r));
	else if (n.localName === "hyperlink") {
		let e = n.getAttributeNS(R, "id") || n.getAttribute("r:id") || "", a = (e ? t.rel(r, e) : void 0)?.target || "";
		for (let o of V(n, "r")) {
			let n = Se(o, t, r);
			n.hyperlink = {
				rId: e,
				url: a
			}, i.runs.push(n);
		}
	} else n.localName === "pPr" || n.localName === "bookmarkStart" || n.localName === "bookmarkEnd" || n.localName === "proofErr" || i.runs.push({ raw: K(n) });
	return i;
}
function Ee(e, r, i) {
	let a = H(e, "tblGrid"), o = a ? V(a, "gridCol").map((e) => Math.round(+(U(e, "w") || 1e3) / B)) : [], s = [];
	for (let t of V(e, "tr")) {
		let e = [];
		for (let a of V(t, "tc")) {
			let t = H(a, "tcPr"), o = { paras: [] }, s = t ? H(t, "gridSpan") : void 0;
			s && (o.gridSpan = +(U(s, "val") || 1));
			let c = t ? H(t, "vMerge") : void 0;
			c && (o.vMerge = U(c, "val") === "restart" ? "restart" : "continue");
			let l = t ? H(t, "shd") : void 0;
			l && U(l, "fill") && U(l, "fill") !== "auto" && (o.fill = "#" + U(l, "fill"));
			let u = t ? H(t, "tcW") : void 0;
			u && (o.width = Math.round(+(U(u, "w") || 0) / B));
			for (let e of V(a, "p")) o.paras.push(Te(e, r, i));
			o.paras.length || o.paras.push(n()), e.push(o);
		}
		s.push(e);
	}
	return {
		kind: "tbl",
		id: t("t"),
		rows: s,
		widths: o
	};
}
function De(e) {
	let t = [], n = "Calibri", r = 11;
	if (!e.has("word/styles.xml")) return {
		styles: t,
		defaultFont: n,
		defaultSize: r
	};
	let i = e.xml("word/styles.xml"), a = H(i, "docDefaults"), o = G(a ? H(H(a, "rPrDefault") || a, "rPr") : void 0);
	o?.font && (n = o.font), o?.size && (r = o.size);
	for (let e of V(i, "style")) {
		let n = U(e, "styleId") || "", r = {
			id: n,
			name: U(H(e, "name"), "val") || n
		}, i = H(e, "basedOn");
		i && (r.basedOn = U(i, "val") || void 0), r.rPr = G(H(e, "rPr")), r.pPr = xe(H(e, "pPr")), t.push(r);
	}
	return {
		styles: t,
		defaultFont: n,
		defaultSize: r
	};
}
function Oe(e) {
	let t = {};
	if (!e.has("word/numbering.xml")) return t;
	let n = e.xml("word/numbering.xml"), r = {};
	for (let e of V(n, "abstractNum")) {
		let t = U(e, "abstractNumId") || "", n = V(e, "lvl").find((e) => U(e, "ilvl") === "0") || V(e, "lvl")[0];
		r[t] = n && U(H(n, "numFmt"), "val") || "bullet";
	}
	for (let e of V(n, "num")) {
		let n = +(U(e, "numId") || 0);
		t[n] = r[U(H(e, "abstractNumId"), "val") || ""] || "bullet";
	}
	return t;
}
function ke(e) {
	let t = H(e, "sectPr"), n = r();
	if (!t) return n;
	let i = H(t, "pgSz");
	i && (n.pageW = +(U(i, "w") || n.pageW), n.pageH = +(U(i, "h") || n.pageH));
	let a = H(t, "pgMar");
	return a && (n.margins = {
		top: +(U(a, "top") || 1440),
		right: +(U(a, "right") || 1440),
		bottom: +(U(a, "bottom") || 1440),
		left: +(U(a, "left") || 1440)
	}), n;
}
async function Ae(e) {
	let t = await L.open(e), r = t.main || "word/document.xml", i = t.xml(r), a = H(i, "body") || i, o = [];
	for (let e of a.children) e.localName === "p" ? o.push(Te(e, t, r)) : e.localName === "tbl" ? o.push(Ee(e, t, r)) : e.localName === "sectPr" || o.push({
		...n(),
		raw: K(e)
	});
	o.length || o.push(n());
	let { styles: s, defaultFont: c, defaultSize: l } = De(t);
	return {
		doc: {
			blocks: o,
			section: ke(a),
			styles: s,
			defaultFont: c,
			defaultSize: l,
			numbering: Oe(t),
			format: t.variant("docx")
		},
		pkg: t,
		part: r
	};
}
// ../vword/src/engine/docx-write.ts
var je = 15, Me = 9525, q = (e) => Math.round(e * je), J = "xmlns:w=\"" + j.w + "\" xmlns:r=\"" + j.r + "\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"" + j.a + "\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"";
function Ne(e) {
	if (!e) return "";
	let t = "";
	return e.font && (t += `<w:rFonts w:ascii="${T(e.font)}" w:hAnsi="${T(e.font)}" w:cs="${T(e.font)}"/>`), e.bold && (t += "<w:b/>"), e.italic && (t += "<w:i/>"), e.strike && (t += "<w:strike/>"), e.underline && (t += "<w:u w:val=\"single\"/>"), e.color && (t += `<w:color w:val="${e.color.replace("#", "").toUpperCase()}"/>`), e.size && (t += `<w:sz w:val="${Math.round(e.size * 2)}"/><w:szCs w:val="${Math.round(e.size * 2)}"/>`), e.highlight && (t += `<w:highlight w:val="${T(e.highlight)}"/>`), e.vertAlign && (t += `<w:vertAlign w:val="${e.vertAlign}"/>`), t ? `<w:rPr>${t}</w:rPr>` : "";
}
function Pe(e) {
	if (!e) return "";
	let t = "";
	if (e.pageBreakBefore && (t += "<w:pageBreakBefore/>"), e.style && (t += `<w:pStyle w:val="${T(e.style)}"/>`), e.numId != null && (t += `<w:numPr><w:ilvl w:val="${e.ilvl || 0}"/><w:numId w:val="${e.numId}"/></w:numPr>`), e.spaceBefore != null || e.spaceAfter != null || e.line != null) {
		let n = e.spaceBefore == null ? "" : ` w:before="${q(e.spaceBefore)}"`, r = e.spaceAfter == null ? "" : ` w:after="${q(e.spaceAfter)}"`, i = "";
		e.line != null && (i = e.line > 4 ? ` w:line="${q(e.line)}" w:lineRule="exact"` : ` w:line="${Math.round(e.line * 240)}" w:lineRule="auto"`), t += `<w:spacing${n}${r}${i}/>`;
	}
	if (e.indent != null || e.hanging != null) {
		let n = e.indent == null ? "" : ` w:left="${q(e.indent)}"`, r = e.hanging == null ? "" : ` w:hanging="${q(e.hanging)}"`;
		t += `<w:ind${n}${r}/>`;
	}
	return e.align && (t += `<w:jc w:val="${e.align === "justify" ? "both" : e.align}"/>`), t ? `<w:pPr>${t}</w:pPr>` : "";
}
function Fe(e, t, n) {
	if (n && e.pkg.rel(e.part, n)) return n;
	let r = e.pkg.relsOf(e.part).find((e) => e.target === t && e.mode === "External");
	if (r) return r.id;
	let i = t;
	if (e.linkIds.has(i)) return e.linkIds.get(i);
	let a = e.pkg.addRel(e.part, j.rel + "/hyperlink", t, "External");
	return e.linkIds.set(i, a), a;
}
async function Ie(e, t) {
	if (t.part && e.pkg.has(t.part) && t.rId && e.pkg.rel(e.part, t.rId)) return t.rId;
	let n = await fetch(t.url), r = new Uint8Array(await n.arrayBuffer()), i = (t.part.split(".").pop() || "png").toLowerCase(), a = e.pkg.nextMediaName(i === "jpg" ? "jpeg" : i), o = (e.part.startsWith("word/") ? "word/media/" : "media/") + a;
	e.pkg.addPart(o, r);
	let s = e.pkg.addRel(e.part, j.rel + "/image", "media/" + a);
	return t.part = o, t.rId = s, s;
}
async function Le(e, t) {
	if (t.raw != null) return t.raw;
	let n = Ne(t.rPr);
	if (t.image) {
		let r = await Ie(e, t.image), i = Math.round(t.image.w * Me), a = Math.round(t.image.h * Me), o = ++e.picId;
		return `<w:r>${n}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${i}" cy="${a}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${o}" name="Picture ${o}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${o}" name="Picture ${o}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${r}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${i}" cy="${a}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
	}
	let r = "";
	if (t.br && (r += "<w:br/>"), t.tab && (r += "<w:tab/>"), t.text != null) {
		let e = /^\s|\s$/.test(t.text) ? " xml:space=\"preserve\"" : "";
		r += `<w:t${e}>${ge(t.text)}</w:t>`;
	}
	r ||= "<w:t xml:space=\"preserve\"></w:t>";
	let i = `<w:r>${n}${r}</w:r>`;
	return t.hyperlink ? `<w:hyperlink r:id="${Fe(e, t.hyperlink.url, t.hyperlink.rId)}">${i}</w:hyperlink>` : i;
}
async function Re(e, t) {
	if (t.raw != null && !t.runs.length) return t.raw;
	let n = [];
	for (let r of t.runs) n.push(await Le(e, r));
	return `<w:p>${Pe(t.pPr)}${n.join("")}</w:p>`;
}
async function ze(e, t) {
	let n = t.widths.map((e) => `<w:gridCol w:w="${q(e)}"/>`).join(""), r = "<w:tblBorders>" + [
		"top",
		"left",
		"bottom",
		"right",
		"insideH",
		"insideV"
	].map((e) => `<w:${e} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join("") + "</w:tblBorders>", i = [];
	for (let n of t.rows) {
		let t = [];
		for (let r of n) {
			let n = `<w:tcW w:w="${q(r.width || 100)}" w:type="dxa"/>`;
			r.gridSpan && r.gridSpan > 1 && (n += `<w:gridSpan w:val="${r.gridSpan}"/>`), r.vMerge && (n += r.vMerge === "restart" ? "<w:vMerge w:val=\"restart\"/>" : "<w:vMerge/>"), r.fill && (n += `<w:shd w:val="clear" w:color="auto" w:fill="${r.fill.replace("#", "").toUpperCase()}"/>`);
			let i = [];
			for (let t of r.paras) i.push(await Re(e, t));
			t.push(`<w:tc><w:tcPr>${n}</w:tcPr>${i.join("")}</w:tc>`);
		}
		i.push(`<w:tr>${t.join("")}</w:tr>`);
	}
	return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${r}</w:tblPr><w:tblGrid>${n}</w:tblGrid>${i.join("")}</w:tbl>`;
}
function Be(e) {
	let t = e.section || r(), n = t.margins;
	return `<w:sectPr><w:pgSz w:w="${t.pageW}" w:h="${t.pageH}"/><w:pgMar w:top="${n.top}" w:right="${n.right}" w:bottom="${n.bottom}" w:left="${n.left}" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>`;
}
async function Ve(e, t) {
	let n = [];
	for (let r of t.blocks) a(r) ? n.push(await Re(e, r)) : o(r) && n.push(await ze(e, r));
	return n.join("") + Be(t);
}
var He = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", Ue = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
function We(e) {
	let t = e.styles.map((e) => {
		let t = `<w:name w:val="${T(e.name)}"/>`;
		return e.basedOn && (t += `<w:basedOn w:val="${T(e.basedOn)}"/>`), t += "<w:qFormat/>", e.pPr && (t += Pe(e.pPr).replace("<w:pPr>", "<w:pPr>").replace("</w:pPr>", "</w:pPr>")), e.rPr && (t += Ne(e.rPr)), `<w:style w:type="paragraph" w:styleId="${T(e.id)}">${t}</w:style>`;
	}).join("");
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:styles ${J}><w:docDefaults><w:rPrDefault>${`<w:rPr><w:rFonts w:ascii="${T(e.defaultFont)}" w:hAnsi="${T(e.defaultFont)}"/><w:sz w:val="${Math.round(e.defaultSize * 2)}"/></w:rPr>`}</w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` + t + "</w:styles>";
}
var Ge = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r
<w:settings ${J}><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/></w:settings>`;
async function Ke(e, t = {}) {
	let n = M[t.format || e.format || "docx"] || M.docx, r = t.part || t.pkg?.main || "word/document.xml", i = t.pkg || be(r, "", n), a = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:document ${J}><w:body>${await Ve({
		pkg: i,
		part: r,
		picId: 100,
		linkIds: /* @__PURE__ */ new Map(),
		newImages: []
	}, e)}</w:body></w:document>`;
	return i.setPart(r, a), t.pkg ? i.has("word/styles.xml") || (i.addPart("word/styles.xml", We(e), He), i.addRel(r, j.rel + "/styles", "styles.xml")) : (i.addPart("word/styles.xml", We(e), He), i.addRel(r, j.rel + "/styles", "styles.xml"), i.addPart("word/settings.xml", Ge, Ue), i.addRel(r, j.rel + "/settings", "settings.xml")), i.save(n.mime);
}
// ../vword/src/engine/render.ts
var qe = {
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
}, Je = (e) => e.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
function Y(e, t) {
	let n = [], r = e?.font || t.defaultFont, i = e?.size || t.defaultSize;
	n.push(`font-family:${r},Carlito,Calibri,sans-serif`), n.push(`font-size:${i}pt`), e?.bold && n.push("font-weight:bold"), e?.italic && n.push("font-style:italic");
	let a = [];
	return e?.underline && a.push("underline"), e?.strike && a.push("line-through"), a.length && n.push("text-decoration:" + a.join(" ")), e?.color && n.push("color:" + e.color), e?.highlight && n.push("background:" + e.highlight), e?.vertAlign === "superscript" && n.push("vertical-align:super;font-size:smaller"), e?.vertAlign === "subscript" && n.push("vertical-align:sub;font-size:smaller"), n.join(";");
}
function Ye(e, t) {
	let n = [], r = e.pPr?.style ? t.styles.find((t) => t.id === e.pPr.style) : void 0, i = e.pPr?.align || r?.pPr?.align;
	i && n.push("text-align:" + i);
	let a = e.pPr?.spaceBefore ?? r?.pPr?.spaceBefore, o = e.pPr?.spaceAfter ?? r?.pPr?.spaceAfter;
	a != null && n.push("margin-top:" + a + "px"), o != null && n.push("margin-bottom:" + o + "px");
	let s = e.pPr?.indent ?? 0;
	s && n.push("padding-left:" + s + "px"), e.pPr?.line && n.push(e.pPr.line > 4 ? "line-height:" + e.pPr.line + "px" : "line-height:" + e.pPr.line);
	let c = r?.rPr;
	return c && n.push(Y(c, t)), n.join(";");
}
function Xe(e, t) {
	if (e.raw != null) return `<span class="raw" contenteditable="false" data-raw="${encodeURIComponent(e.raw)}">[object]</span>`;
	let n = Y(e.rPr, t);
	if (e.image) return `<img class="img" src="${e.image.url}" width="${e.image.w}" height="${e.image.h}" data-part="${e.image.part}" data-rid="${e.image.rId}" data-w="${e.image.w}" data-h="${e.image.h}" style="${n}">`;
	if (e.br) return "<br>";
	if (e.tab) return `<span class="tb" style="${n}">\t</span>`;
	let r = `<span style="${n}">${Je(e.text ?? "") || ""}</span>`;
	return e.hyperlink ? `<a href="${Je(e.hyperlink.url)}" data-rid="${e.hyperlink.rId}" class="lnk">${r}</a>` : r;
}
function X(e, t) {
	if (a(e)) {
		let n = e.runs.map((e) => Xe(e, t)).join(""), r = Ze(e, t);
		return `<div class="p" data-id="${e.id}" style="${Ye(e, t)}">${r}${n || "<br>"}</div>`;
	}
	let n = e, r = n.widths.map((e) => `<col style="width:${e}px">`).join(""), i = n.rows.map((e) => `<tr>${e.map((e) => {
		let n = e.gridSpan && e.gridSpan > 1 ? ` colspan="${e.gridSpan}"` : "", r = e.vMerge === "continue" ? " data-vmerge=\"continue\"" : e.vMerge === "restart" ? " data-vmerge=\"restart\"" : "", i = e.fill ? `background:${e.fill};` : "", a = e.paras.map((e) => X(e, t)).join("");
		return `<td${n}${r} data-w="${e.width || 0}" style="${i}border:1px solid #888;padding:2px 4px;vertical-align:top">${a}</td>`;
	}).join("")}</tr>`).join("");
	return `<table class="tbl" data-id="${n.id}" style="border-collapse:collapse;table-layout:fixed"><colgroup>${r}</colgroup>${i}</table>`;
}
function Ze(e, t) {
	return e.pPr?.numId == null ? "" : `<span class="lm" contenteditable="false" data-num="${e.pPr.numId}" data-ilvl="${e.pPr.ilvl || 0}"></span>`;
}
function Z(e) {
	if (!e) return;
	if (e.startsWith("#")) return e.length === 7 ? e : void 0;
	let t = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(e);
	if (!t) return;
	let n = (e) => (+e).toString(16).padStart(2, "0");
	return "#" + n(t[1]) + n(t[2]) + n(t[3]);
}
function Qe(e, t) {
	let n = { ...e }, r = t.tagName.toLowerCase();
	(r === "b" || r === "strong") && (n.bold = !0), (r === "i" || r === "em") && (n.italic = !0), r === "u" && (n.underline = !0), (r === "s" || r === "strike" || r === "del") && (n.strike = !0);
	let i = t.style;
	i.fontWeight && (i.fontWeight === "bold" || +i.fontWeight >= 600) && (n.bold = !0), i.fontWeight && (i.fontWeight === "normal" || +i.fontWeight < 600) && (n.bold = !1), i.fontStyle === "italic" && (n.italic = !0), i.fontStyle === "normal" && (n.italic = !1);
	let a = i.textDecorationLine || i.textDecoration || "";
	if (a.includes("underline") && (n.underline = !0), a.includes("line-through") && (n.strike = !0), i.fontFamily && (n.font = i.fontFamily.split(",")[0].replace(/["']/g, "").trim()), i.fontSize) {
		let e = parseFloat(i.fontSize);
		i.fontSize.endsWith("pt") ? n.size = e : isNaN(e) || (n.size = Math.round(e * .75 * 2) / 2);
	}
	let o = Z(i.color);
	o && (n.color = o);
	let s = Z(i.backgroundColor);
	return s && s !== "#ffffff" && (n.highlight = qe[s.toLowerCase()] || s), i.verticalAlign === "super" && (n.vertAlign = "superscript"), i.verticalAlign === "sub" && (n.vertAlign = "subscript"), n;
}
function $e(e, t, n, r) {
	if (e.nodeType === 3) {
		let i = e.nodeValue || "";
		i && r.push({
			text: i,
			rPr: Q(t),
			hyperlink: n
		});
		return;
	}
	if (e.nodeType !== 1) return;
	let i = e, a = i.tagName.toLowerCase();
	if (a === "br") {
		r.push({
			br: !0,
			rPr: Q(t)
		});
		return;
	}
	if (a === "img") {
		r.push({
			image: {
				url: i.getAttribute("src") || "",
				part: i.getAttribute("data-part") || "",
				rId: i.getAttribute("data-rid") || "",
				w: +(i.getAttribute("data-w") || i.getAttribute("width") || 100),
				h: +(i.getAttribute("data-h") || i.getAttribute("height") || 100)
			},
			rPr: Q(t)
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
			rPr: Q(t)
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
	}) : o = Qe(t, i);
	for (let e of Array.from(i.childNodes)) $e(e, o, s, r);
}
function Q(e) {
	let t = {};
	return e.bold && (t.bold = !0), e.italic && (t.italic = !0), e.underline && (t.underline = !0), e.strike && (t.strike = !0), e.font && (t.font = e.font), e.size && (t.size = e.size), e.color && (t.color = e.color), e.highlight && (t.highlight = e.highlight), e.vertAlign && (t.vertAlign = e.vertAlign), Object.keys(t).length ? t : void 0;
}
function $(e, t) {
	let r = [];
	for (let t of Array.from(e.childNodes)) $e(t, {}, void 0, r);
	let i = t ? {
		...t,
		runs: r
	} : n();
	i.id = e.getAttribute("data-id") || i.id;
	let a = e.style.textAlign;
	return a && (i.pPr = {
		...i.pPr,
		align: a
	}), i;
}
function et(e, r, i) {
	let a = e.getAttribute("data-id") || void 0;
	if (e.tagName.toLowerCase() === "table") {
		let r = i && i.kind === "tbl" ? i : void 0, o = Array.from(e.querySelectorAll("col")).map((e) => parseInt(e.style.width) || 100), s = Array.from(e.querySelectorAll(":scope > tbody > tr, :scope > tr")).map((e) => Array.from(e.querySelectorAll(":scope > td, :scope > th")).map((e) => {
			let t = e, r = Array.from(t.querySelectorAll(":scope > .p")).map((e) => $(e));
			r.length || r.push(n());
			let i = { paras: r }, a = t.getAttribute("colspan");
			a && +a > 1 && (i.gridSpan = +a);
			let o = t.getAttribute("data-vmerge");
			(o === "restart" || o === "continue") && (i.vMerge = o);
			let s = Z(t.style.backgroundColor);
			s && (i.fill = s);
			let c = +(t.getAttribute("data-w") || 0);
			return c && (i.width = c), i;
		}));
		return {
			kind: "tbl",
			id: a || t("t"),
			rows: s,
			widths: o.length ? o : r?.widths || [100]
		};
	}
	return $(e, i && i.kind === "p" ? i : void 0);
}
// ../vword/src/engine/layout.ts
function tt(e, t, n, r) {
	r.style.width = t + "px", r.innerHTML = "";
	let i = [[]], a = 0;
	for (let t of e.blocks) {
		let o = document.createElement("div");
		o.innerHTML = X(t, e);
		let s = o.firstElementChild;
		r.appendChild(s);
		let c = s.offsetHeight;
		r.removeChild(s);
		let l = t.kind === "p" && t.pPr?.pageBreakBefore;
		(a + c > n || l) && i[i.length - 1].length ? (i.push([t]), a = c) : (i[i.length - 1].push(t), a += c);
	}
	return i;
}
function nt(e, t) {
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
			n.set(o, e), r.textContent = (a === "lowerLetter" ? String.fromCharCode(96 + e) : a === "upperLetter" ? String.fromCharCode(64 + e) : a === "lowerRoman" ? rt(e) : String(e)) + ". ";
		}
	}
}
function rt(e) {
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
function it(e) {
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
function at(e, t) {
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
export { c as _, X as a, Ae as c, r as d, l as f, i as g, o as h, at as i, s as l, a as m, tt as n, et as o, n as p, nt as r, Ke as s, it as t, d as u, t as v, u as y };
