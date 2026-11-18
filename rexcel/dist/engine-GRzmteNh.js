// ../vexcel/src/engine/model.ts
var e = (e) => {
	let t = "";
	for (e++; e > 0;) {
		let n = (e - 1) % 26;
		t = String.fromCharCode(65 + n) + t, e = Math.floor((e - 1) / 26);
	}
	return t;
}, t = (e) => {
	let t = 0;
	for (let n of e.toUpperCase()) t = t * 26 + (n.charCodeAt(0) - 64);
	return t - 1;
}, n = (t, n) => e(n) + (t + 1);
function r(e) {
	let n = /^([A-Za-z]+)(\d+)$/.exec(e);
	return n ? {
		r: n[2] - 1,
		c: t(n[1])
	} : null;
}
var i = (e) => ({
	name: e,
	cells: /* @__PURE__ */ new Map(),
	merges: [],
	colWidths: /* @__PURE__ */ new Map(),
	rowHeights: /* @__PURE__ */ new Map(),
	defaultColWidth: 64,
	defaultRowHeight: 17
}), a = () => ({
	sheets: [i("Sheet1")],
	styles: [{}],
	active: 0
});
function o(e, t) {
	let n = JSON.stringify(t), r = e.styles.findIndex((e) => JSON.stringify(e) === n);
	return r >= 0 ? r : (e.styles.push(t), e.styles.length - 1);
}
function s(e, t, r) {
	return e.cells.get(n(t, r));
}
function c(e, t, r, i) {
	let a = n(t, r);
	!i || i.v == null && !i.f && i.s == null ? e.cells.delete(a) : e.cells.set(a, i);
}
function l(e, t, n) {
	if (!e) return "";
	if (e.t === "e") return String(e.v);
	let r = e.s == null ? void 0 : t.styles[e.s];
	return n(e.v, r?.numFmt);
}
var u = (e) => ({
	sheets: e.sheets.map((e) => ({
		...e,
		cells: new Map(e.cells),
		merges: e.merges.map((e) => ({ ...e })),
		colWidths: new Map(e.colWidths),
		rowHeights: new Map(e.rowHeights)
	})),
	styles: e.styles.map((e) => ({ ...e })),
	active: e.active
}), d = new TextEncoder(), f = new TextDecoder("utf-8"), p = (() => {
	let e = /* @__PURE__ */ new Uint32Array(256);
	for (let t = 0; t < 256; t++) {
		let n = t;
		for (let e = 0; e < 8; e++) n = n & 1 ? 3988292384 ^ n >>> 1 : n >>> 1;
		e[t] = n >>> 0;
	}
	return e;
})();
function m(e) {
	let t = 4294967295;
	for (let n = 0; n < e.length; n++) t = p[(t ^ e[n]) & 255] ^ t >>> 8;
	return (t ^ 4294967295) >>> 0;
}
var h = [
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
], g = [
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
], _ = [
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
], ee = [
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
], te = [
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
function ne() {
	if (y) return;
	let e = /* @__PURE__ */ new Uint8Array(288);
	for (let t = 0; t < 144; t++) e[t] = 8;
	for (let t = 144; t < 256; t++) e[t] = 9;
	for (let t = 256; t < 280; t++) e[t] = 7;
	for (let t = 280; t < 288; t++) e[t] = 8;
	y = v(e, 288), b = v((/* @__PURE__ */ new Uint8Array(30)).fill(5), 30);
}
function re(e, t) {
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
				let e = h[n] + c(g[n]), r = l(t), i = _[r] + c(ee[r]);
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
		} else if (t === 1) ne(), u(y, b);
		else if (t === 2) {
			let e = c(5) + 257, t = c(5) + 1, n = c(4) + 4, r = /* @__PURE__ */ new Uint8Array(19);
			for (let e = 0; e < n; e++) r[te[e]] = c(3);
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
async function ie(e, t) {
	let n = new Response(new Blob([e]).stream().pipeThrough(t));
	return new Uint8Array(await n.arrayBuffer());
}
async function ae(e, t) {
	if (typeof DecompressionStream == "function") try {
		return await ie(e, new DecompressionStream("deflate-raw"));
	} catch {}
	return re(e, t);
}
async function x(e) {
	if (typeof CompressionStream == "function") try {
		return await ie(e, new CompressionStream("deflate-raw"));
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
			return f.decode(await this.bytes());
		}
	};
}
var C = (e, t) => t & 2048 ? f.decode(e) : Array.from(e, (e) => String.fromCharCode(e)).join("");
function w(e, t) {
	if (t.name.endsWith("/")) return;
	e.set(t.name, t);
	let n = t.name.replace(/\\/g, "/");
	n !== t.name && !e.has(n) && e.set(n, t);
}
function oe(e, t) {
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
function T(e, t) {
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
async function E(e) {
	let t = e instanceof Uint8Array ? e : new Uint8Array(e), n = new DataView(t.buffer, t.byteOffset, t.byteLength);
	if (t.length < 22) throw Error("This file is not a valid Office Open XML (ZIP) package.");
	if (n.getUint32(0, !0) !== 67324752 && n.getUint32(0, !0) === 3759263696) {
		let e = new TextDecoder("utf-16le").decode(t.subarray(0, Math.min(t.length, 1 << 20))), n = /EncryptedPackage/.test(e), r = /* @__PURE__ */ Error(n ? "This file is password-protected." : "This is a binary Office 97–2003 file, not an Office Open XML package.");
		throw r.code = "ole", r.encrypted = n, r.bytes = t, r;
	}
	let r = null;
	try {
		r = oe(t, n);
	} catch {
		r = null;
	}
	if (!r) {
		try {
			r = T(t, n);
		} catch {
			r = null;
		}
		if (!r || !r.size) throw Error(n.getUint32(0, !0) === 67324752 ? "This file is damaged and could not be repaired." : "This file is not a valid Office Open XML (ZIP) package.");
		r.repaired = !0;
	}
	return r;
}
function se(e) {
	return {
		time: e.getHours() << 11 | e.getMinutes() << 5 | Math.floor(e.getSeconds() / 2),
		date: e.getFullYear() - 1980 << 9 | e.getMonth() + 1 << 5 | e.getDate()
	};
}
async function ce(e, t) {
	let n = [], r = [], i = 0, { time: a, date: o } = se(/* @__PURE__ */ new Date());
	for (let t of e) {
		let e = d.encode(t.name), s = typeof t.data == "string" ? d.encode(t.data) : t.data, c = m(s), l = null;
		!t.store && s.length > 64 && (l = await x(s));
		let u = l && l.length < s.length ? 8 : 0, f = u === 8 ? l : s, p = new Uint8Array(30 + e.length), h = new DataView(p.buffer);
		h.setUint32(0, 67324752, !0), h.setUint16(4, 20, !0), h.setUint16(6, 2048, !0), h.setUint16(8, u, !0), h.setUint16(10, a, !0), h.setUint16(12, o, !0), h.setUint32(14, c, !0), h.setUint32(18, f.length, !0), h.setUint32(22, s.length, !0), h.setUint16(26, e.length, !0), h.setUint16(28, 0, !0), p.set(e, 30), n.push(p, f);
		let g = new Uint8Array(46 + e.length), _ = new DataView(g.buffer);
		_.setUint32(0, 33639248, !0), _.setUint16(4, 20, !0), _.setUint16(6, 20, !0), _.setUint16(8, 2048, !0), _.setUint16(10, u, !0), _.setUint16(12, a, !0), _.setUint16(14, o, !0), _.setUint32(16, c, !0), _.setUint32(20, f.length, !0), _.setUint32(24, s.length, !0), _.setUint16(28, e.length, !0), _.setUint32(42, i, !0), g.set(e, 46), r.push(g), i += p.length + f.length;
	}
	let s = r.reduce((e, t) => e + t.length, 0), c = /* @__PURE__ */ new Uint8Array(22), l = new DataView(c.buffer);
	return l.setUint32(0, 101010256, !0), l.setUint16(8, e.length, !0), l.setUint16(10, e.length, !0), l.setUint32(12, s, !0), l.setUint32(16, i, !0), new Blob([
		...n,
		...r,
		c
	], { type: t || "application/zip" });
}
// ../voffice-shared/src/core/xml.ts
var le = {
	lt: "<",
	gt: ">",
	amp: "&",
	quot: "\"",
	apos: "'"
};
function ue(e) {
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
		return le[t] == null ? e : le[t];
	});
}
var D = (e) => (e = String(e).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"), /[\t\n\r]/.test(e) ? e.replace(/\t/g, "&#9;").replace(/\n/g, "&#10;").replace(/\r/g, "&#13;") : e), de = /* @__PURE__ */ RegExp("[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F\\uFFFE\\uFFFF]", "g"), fe = /* @__PURE__ */ RegExp("[\\x00-\\x08\\x0B-\\x1F\\uFFFE\\uFFFF]", "g"), pe = (e, t) => String(e).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/_x([0-9A-Fa-f]{4})_/g, "_x005F_x$1_").replace(t ? de : fe, (e) => "_x" + e.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0") + "_"), me = class {
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
}, he = class e {
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
		this.children = [], this.childNodes = [new me(String(e), this)];
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
		if (this.attrs && (n.attrs = { ...this.attrs }), t) for (let e of this.childNodes) n.appendChild(e.nodeType === 3 ? new me(e.data, n) : e.cloneNode(!0));
		return n;
	}
}, ge = /([^\s=\/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
function _e(e, t) {
	ge.lastIndex = 0;
	let n;
	for (; n = ge.exec(e);) {
		let e = t.attrs ??= Object.create(null), r = n[2] === void 0 ? n[3] : n[2];
		e[n[1]] = r.indexOf("&") < 0 ? r : ue(r);
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
function ye(e, t) {
	e && e.charCodeAt(0) === 65279 && (e = e.slice(1));
	let n = !!t?.keepWhitespace, r = !!t?.strict, i = new he("#document", null), a = i, o = 0, s = e.length, c = (e) => {
		a !== i && (n || !/^\s*$/.test(e) || ve.has(a.localName)) && a.childNodes.push(new me(ue(e), a));
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
				a !== i && a.childNodes.push(new me(r, a)), o = n < 0 ? s : n + 3;
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
			let c = e.slice(t + 1, i), l = e.charCodeAt(n - 1) === 47, u = new he(c, a), d = l ? n - 1 : n;
			i < d && _e(e.slice(i, d), u), a.childNodes.push(u), a.children.push(u), l || (a = u), o = n + 1;
		}
	}
	let l = i.children[0];
	if (!l) throw Error("XML parse error: no root element");
	if (r && (a !== i || i.children.length !== 1)) throw Error("XML parse error: incomplete document");
	return l;
}
function be(e) {
	return e.length >= 2 && e[0] === 255 && e[1] === 254 ? new TextDecoder("utf-16le").decode(e.subarray(2)) : e.length >= 2 && e[0] === 254 && e[1] === 255 ? new TextDecoder("utf-16be").decode(e.subarray(2)) : e.length >= 4 && e[0] === 0 && e[1] === 60 && e[2] === 0 && e[3] === 63 ? new TextDecoder("utf-16be").decode(e) : e.length >= 4 && e[0] === 60 && e[1] === 0 && e[2] === 63 && e[3] === 0 ? new TextDecoder("utf-16le").decode(e) : new TextDecoder("utf-8").decode(e);
}
// ../voffice-shared/src/core/opc.ts
var O = Object.freeze({
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
}), xe = Object.freeze(Object.fromEntries([
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
}))), Se = "[Content_Types].xml", Ce = (e) => {
	let t = e.lastIndexOf("/");
	return (t < 0 ? "" : e.slice(0, t + 1)) + "_rels/" + (t < 0 ? e : e.slice(t + 1)) + ".rels";
};
function we(e, t) {
	if (t.startsWith("/")) return t.slice(1);
	let n = ((e.includes("/") ? e.slice(0, e.lastIndexOf("/") + 1) : "") + t).split("/"), r = [];
	for (let e of n) e !== "." && e !== "" && (e === ".." ? r.pop() : r.push(e));
	return r.join("/");
}
var Te = (e) => e.replace(O.strictRel + "/", O.rel + "/"), Ee = class e {
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
		let n = await E(t), r = new e();
		r.repaired = !!n.repaired;
		for (let [e, t] of n) r.parts.set(e, await t.bytes());
		let i = r.parts.get(Se);
		if (!i) throw Error("This package has no [Content_Types].xml part.");
		let a = ye(be(i));
		for (let e of a.children) e.localName === "Default" ? r.defaults.set(e.getAttribute("Extension") || "", e.getAttribute("ContentType") || "") : e.localName === "Override" && r.overrides.set((e.getAttribute("PartName") || "").replace(/^\//, ""), e.getAttribute("ContentType") || "");
		let o = r.relsOf("").find((e) => Te(e.type) === O.rel + "/officeDocument");
		return o && (r.main = we("", o.target)), r;
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
			t = be(n), this.textCache.set(e, t);
		}
		return t;
	}
	xml(e) {
		let t = this.xmlCache.get(e);
		return t || (t = ye(this.text(e)), this.xmlCache.set(e, t)), t;
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
		let r = this.parts.get(Ce(e));
		if (r) {
			let e = ye(be(r));
			for (let t of e.children) t.localName === "Relationship" && n.push({
				id: t.getAttribute("Id") || "",
				type: Te(t.getAttribute("Type") || ""),
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
		return this.relsOf(e).filter((e) => e.type === t && e.mode !== "External").map((t) => we(e, t.target));
	}
	setPart(e, t) {
		this.parts.set(e, typeof t == "string" ? new TextEncoder().encode(t) : t), this.edited.add(e), this.textCache.delete(e), this.xmlCache.delete(e);
	}
	addPart(e, t, n) {
		if (this.setPart(e, t), n) this.overrides.set(e, n), this.ctDirty = !0;
		else {
			let t = e.includes(".") ? e.slice(e.lastIndexOf(".") + 1) : "";
			t && !this.defaults.has(t) && (this.defaults.set(t, De(t)), this.ctDirty = !0);
		}
	}
	removePart(e) {
		this.parts.delete(e), this.edited.delete(e), this.textCache.delete(e), this.xmlCache.delete(e), this.overrides.delete(e) && (this.ctDirty = !0);
		let t = Ce(e);
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
		return Object.values(xe).find((e) => e.contentType === t)?.ext || e;
	}
	writeContentTypes() {
		if (!this.ctDirty) return;
		let e = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", "<Types xmlns=\"" + O.ct + "\">"];
		for (let [t, n] of [...this.defaults].sort()) e.push(`<Default Extension="${D(t)}" ContentType="${D(n)}"/>`);
		for (let [t, n] of [...this.overrides].sort()) e.push(`<Override PartName="/${D(t)}" ContentType="${D(n)}"/>`);
		e.push("</Types>"), this.parts.set(Se, new TextEncoder().encode(e.join("")));
	}
	writeRels() {
		for (let e of this.relsDirty) {
			let t = this.relsCache.get(e) || [], n = e, r = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", "<Relationships xmlns=\"" + O.pkg + "\">"];
			for (let e of t) r.push(`<Relationship Id="${D(e.id)}" Type="${D(e.type)}" Target="${D(e.target)}"${e.mode ? ` TargetMode="${D(e.mode)}"` : ""}/>`);
			r.push("</Relationships>");
			let i = Ce(n);
			this.parts.set(i, new TextEncoder().encode(r.join(""))), this.overrides.has(i) || (this.defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml"), this.ctDirty = !0);
		}
	}
	async save(e) {
		return this.writeRels(), this.writeContentTypes(), ce([...this.parts].map(([e, t]) => ({
			name: e,
			data: t
		})), e || "application/zip");
	}
};
function De(e) {
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
function Oe(e, t, n, r = {}) {
	let i = new Ee();
	i.main = e, i.defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml"), i.defaults.set("xml", "application/xml");
	for (let [e, t] of Object.entries(r)) i.defaults.set(e, t);
	return i.overrides.set(e, n.contentType), i.setPart(e, t), i.addRel("", O.rel + "/officeDocument", e), i;
}
// ../vexcel/src/engine/xlsx-read.ts
O.s;
var ke = O.r, k = (e, t) => e.children.filter((e) => e.localName === t), A = (e, t) => e.children.find((e) => e.localName === t), j = (e, t) => e ? e.getAttribute(t) : null, Ae = {
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
function je(e) {
	let t = A(e, "t");
	return t ? t.textContent : k(e, "r").map((e) => A(e, "t")?.textContent || "").join("");
}
function Me(e) {
	let t = {
		fonts: [],
		fills: [],
		borders: [],
		xfs: [],
		numFmts: /* @__PURE__ */ new Map()
	};
	if (!e.has("xl/styles.xml")) return t;
	let n = e.xml("xl/styles.xml"), r = A(n, "numFmts");
	if (r) for (let e of k(r, "numFmt")) t.numFmts.set(+(j(e, "numFmtId") || 0), j(e, "formatCode") || "General");
	let i = A(n, "fonts");
	i && (t.fonts = k(i, "font"));
	let a = A(n, "fills");
	a && (t.fills = k(a, "fill"));
	let o = A(n, "borders");
	o && (t.borders = k(o, "border"));
	let s = A(n, "cellXfs");
	return s && (t.xfs = k(s, "xf")), t;
}
function Ne(e) {
	if (!e) return;
	let t = j(e, "rgb");
	if (t && t.length === 8) return ("#" + t.slice(2)).toLowerCase();
	if (t && t.length === 6) return ("#" + t).toLowerCase();
}
function Pe(e, t) {
	let n = {};
	if (!e) return n;
	let r = t.fonts[+(j(e, "fontId") || 0)];
	if (r) {
		let e = {}, t = A(r, "sz");
		t && (e.size = +(j(t, "val") || 11));
		let i = A(r, "name");
		i && (e.name = j(i, "val") || void 0), A(r, "b") && (e.bold = !0), A(r, "i") && (e.italic = !0), A(r, "u") && (e.underline = !0);
		let a = Ne(A(r, "color"));
		a && (e.color = a), Object.keys(e).length && (n.font = e);
	}
	let i = t.fills[+(j(e, "fillId") || 0)];
	if (i) {
		let e = A(i, "patternFill"), t = Ne(A(e || i, "fgColor"));
		t && j(e, "patternType") !== "none" && (n.fill = t);
	}
	let a = t.borders[+(j(e, "borderId") || 0)];
	if (a) {
		let e = {};
		for (let t of [
			"left",
			"right",
			"top",
			"bottom"
		]) {
			let n = A(a, t);
			n && j(n, "style") && (e[t] = Ne(A(n, "color")) || "#000000");
		}
		Object.keys(e).length && (n.border = e);
	}
	let o = A(e, "alignment");
	if (o) {
		let e = {}, t = j(o, "horizontal");
		t && (e.h = t === "general" ? void 0 : t);
		let r = j(o, "vertical");
		r && (e.v = r === "center" ? "middle" : r), j(o, "wrapText") === "1" && (e.wrap = !0), Object.keys(e).length && (n.align = e);
	}
	let s = +(j(e, "numFmtId") || 0), c = t.numFmts.get(s) || Ae[s];
	return c && c !== "General" && (n.numFmt = c), n;
}
async function Fe(e) {
	let t = await Ee.open(e), n = t.xml("xl/workbook.xml"), r = t.relsOf("xl/workbook.xml"), a = [];
	if (t.has("xl/sharedStrings.xml")) {
		let e = t.xml("xl/sharedStrings.xml");
		for (let t of k(e, "si")) a.push(je(t));
	}
	let o = Me(t), s = {
		sheets: [],
		styles: [{}],
		active: 0
	}, c = [], l = A(n, "sheets"), u = l ? k(l, "sheet") : [];
	for (let e of u) {
		let n = j(e, "name") || "Sheet", l = e.getAttributeNS(ke, "id") || j(e, "r:id") || "", u = r.find((e) => e.id === l), d = u ? u.target.startsWith("/") ? u.target.slice(1) : "xl/" + u.target.replace(/^\.\//, "") : "";
		c.push(d);
		let f = i(n);
		d && t.has(d) && Ie(t, d, f, a, o, s), s.sheets.push(f);
	}
	s.sheets.length || (s.sheets.push(i("Sheet1")), c.push("xl/worksheets/sheet1.xml"));
	let d = j(A(n, "bookViews") && A(A(n, "bookViews"), "workbookView"), "activeTab");
	return d && (s.active = Math.min(+d, s.sheets.length - 1)), {
		wb: s,
		pkg: t,
		sheetParts: c
	};
}
function Ie(e, t, i, a, s, c) {
	let l = e.xml(t), u = A(l, "cols");
	if (u) for (let e of k(u, "col")) {
		let t = +(j(e, "min") || 1), n = +(j(e, "max") || t), r = +(j(e, "width") || 8.43);
		for (let e = t; e <= n && e <= 16384; e++) i.colWidths.set(e - 1, Math.round(r * 7 + 5));
	}
	let d = A(l, "sheetData");
	if (d) for (let e of k(d, "row")) {
		let t = (j(e, "r") || 1) - 1, l = j(e, "ht");
		l && i.rowHeights.set(t, Math.round(+l));
		for (let t of k(e, "c")) {
			let e = r(j(t, "r") || "");
			if (!e) continue;
			let l = j(t, "t") || "n", u = j(t, "s"), d = A(t, "v"), f = A(t, "f"), p = {};
			if (f && (p.f = f.textContent), l === "s") p.v = a[+(d?.textContent || 0)] ?? "", p.t = "s";
			else if (l === "inlineStr") {
				let e = A(t, "is");
				p.v = e ? je(e) : "", p.t = "s";
			} else if (l === "str") p.v = d?.textContent || "", p.t = "s";
			else if (l === "b") p.v = d?.textContent === "1", p.t = "b";
			else if (l === "e") p.v = d?.textContent || "#VALUE!", p.t = "e";
			else {
				let e = d?.textContent;
				p.v = e != null && e !== "" ? +e : null, p.t = "n";
			}
			if (u != null) {
				let e = Pe(s.xfs[+u], s);
				Object.keys(e).length && (p.s = o(c, e));
			}
			(p.v != null || p.f || p.s != null) && i.cells.set(n(e.r, e.c), p);
		}
	}
	let f = A(l, "mergeCells");
	if (f) for (let e of k(f, "mergeCell")) {
		let [t, n] = (j(e, "ref") || "").split(":"), a = r(t), o = r(n || t);
		a && o && i.merges.push({
			r1: a.r,
			c1: a.c,
			r2: o.r,
			c2: o.c
		});
	}
	let p = A(l, "sheetViews"), m = p && A(A(p, "sheetView") || p, "pane");
	if (m) {
		let e = +(j(m, "xSplit") || 0), t = +(j(m, "ySplit") || 0);
		(e || t) && (i.frozen = {
			rows: t,
			cols: e
		});
	}
}
// ../vexcel/src/engine/xlsx-write.ts
var M = `xmlns="${O.s}" xmlns:r="${O.r}"`;
function Le(e) {
	return Math.max(0, Math.round((e - 5) / 7 * 100) / 100);
}
function Re(e, t) {
	let n = e.strings.get(t);
	return n ?? (n = e.strList.length, e.strList.push(t), e.strings.set(t, n)), n;
}
function ze(e, t, r, i) {
	let a = n(t, r), o = i.s ? ` s="${i.s}"` : "";
	return i.f == null ? i.t === "e" ? `<c r="${a}"${o} t="e"><v>${pe(String(i.v))}</v></c>` : i.t === "b" ? `<c r="${a}"${o} t="b"><v>${+!!i.v}</v></c>` : typeof i.v == "string" ? `<c r="${a}"${o} t="s"><v>${Re(e, i.v)}</v></c>` : typeof i.v == "number" ? `<c r="${a}"${o}><v>${i.v}</v></c>` : "" : `<c r="${a}"${o}><f>${pe(i.f)}</f></c>`;
}
function Be(t, i) {
	let a = /* @__PURE__ */ new Map(), o = 0, s = 0;
	for (let [e, t] of i.cells) {
		let n = r(e);
		o = Math.max(o, n.r), s = Math.max(s, n.c);
		let i = a.get(n.r) || [];
		i.push({
			r: n.r,
			c: n.c,
			cell: t
		}), a.set(n.r, i);
	}
	let c = `A1:${e(s)}${o + 1}`, l = "";
	i.colWidths.size && (l = "<cols>" + [...i.colWidths.entries()].sort((e, t) => e[0] - t[0]).map(([e, t]) => `<col min="${e + 1}" max="${e + 1}" width="${Le(t)}" customWidth="1"/>`).join("") + "</cols>");
	let u = [];
	for (let [e, n] of [...a.entries()].sort((e, t) => e[0] - t[0])) {
		n.sort((e, t) => e.c - t.c);
		let r = i.rowHeights.get(e), a = r == null ? "" : ` ht="${r}" customHeight="1"`;
		u.push(`<row r="${e + 1}"${a}>` + n.map((e) => ze(t, e.r, e.c, e.cell)).join("") + "</row>");
	}
	let d = "";
	i.merges.length && (d = `<mergeCells count="${i.merges.length}">` + i.merges.map((e) => `<mergeCell ref="${n(e.r1, e.c1)}:${n(e.r2, e.c2)}"/>`).join("") + "</mergeCells>");
	let f = "";
	if (i.frozen && (i.frozen.rows || i.frozen.cols)) {
		let { rows: e, cols: t } = i.frozen;
		f = `<pane xSplit="${t}" ySplit="${e}" topLeftCell="${n(e, t)}" activePane="bottomRight" state="frozen"/>`;
	}
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet ${M}><dimension ref="${c}"/><sheetViews><sheetView workbookViewId="0">${f}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>${l}<sheetData>${u.join("")}</sheetData>${d}</worksheet>`;
}
function Ve(e) {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<sst ${M} count="${e.strList.length}" uniqueCount="${e.strList.length}">` + e.strList.map((e) => `<si><t xml:space="preserve">${pe(e)}</t></si>`).join("") + "</sst>";
}
function He(e) {
	let t = [], n = ["<fill><patternFill patternType=\"none\"/></fill>", "<fill><patternFill patternType=\"gray125\"/></fill>"], r = ["<border><left/><right/><top/><bottom/><diagonal/></border>"], i = [], a = /* @__PURE__ */ new Map(), o = [], s = /* @__PURE__ */ new Map(), c = /* @__PURE__ */ new Map(), l = /* @__PURE__ */ new Map(), u = 164;
	for (let d of e.styles) {
		let e = 0;
		if (d.font) {
			let n = JSON.stringify(d.font);
			if (!s.has(n)) {
				let e = d.font, r = "";
				e.bold && (r += "<b/>"), e.italic && (r += "<i/>"), e.underline && (r += "<u/>"), r += `<sz val="${e.size || 11}"/>`, e.color && (r += `<color rgb="FF${e.color.replace("#", "").toUpperCase()}"/>`), r += `<name val="${D(e.name || "Calibri")}"/>`, s.set(n, t.length), t.push(`<font>${r}</font>`);
			}
			e = s.get(n);
		}
		let f = 0;
		if (d.fill) {
			let e = d.fill;
			c.has(e) || (c.set(e, n.length), n.push(`<fill><patternFill patternType="solid"><fgColor rgb="FF${e.replace("#", "").toUpperCase()}"/><bgColor indexed="64"/></patternFill></fill>`)), f = c.get(e);
		}
		let p = 0;
		if (d.border) {
			let e = JSON.stringify(d.border);
			if (!l.has(e)) {
				let t = d.border, n = "";
				for (let e of [
					"left",
					"right",
					"top",
					"bottom"
				]) n += t[e] ? `<${e} style="thin"><color rgb="FF${t[e].replace("#", "").toUpperCase()}"/></${e}>` : `<${e}/>`;
				l.set(e, r.length), r.push(`<border>${n}<diagonal/></border>`);
			}
			p = l.get(e);
		}
		let m = 0;
		d.numFmt && d.numFmt !== "General" && (a.has(d.numFmt) || (a.set(d.numFmt, u), i.push(`<numFmt numFmtId="${u}" formatCode="${D(d.numFmt)}"/>`), u++), m = a.get(d.numFmt));
		let h = "";
		if (d.align) {
			let e = d.align;
			h = `<alignment${e.h ? ` horizontal="${e.h}"` : ""}${e.v ? ` vertical="${e.v === "middle" ? "center" : e.v}"` : ""}${e.wrap ? " wrapText=\"1\"" : ""}/>`;
		}
		o.push(`<xf numFmtId="${m}" fontId="${e}" fillId="${f}" borderId="${p}" xfId="0"${m ? " applyNumberFormat=\"1\"" : ""}${e ? " applyFont=\"1\"" : ""}${f ? " applyFill=\"1\"" : ""}${p ? " applyBorder=\"1\"" : ""}${h ? " applyAlignment=\"1\"" : ""}>${h}</xf>`);
	}
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<styleSheet ${M}>` + (i.length ? `<numFmts count="${i.length}">${i.join("")}</numFmts>` : "") + `<fonts count="${t.length || 1}">${t.join("") || "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>"}</fonts><fills count="${n.length}">${n.join("")}</fills><borders count="${r.length}">${r.join("")}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${o.length}">${o.join("")}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
}
var Ue = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", We = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", Ge = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
async function Ke(e, t = {}) {
	let n = xe[t.format || "xlsx"] || xe.xlsx, r = {
		wb: e,
		strings: /* @__PURE__ */ new Map(),
		strList: []
	}, i = e.sheets.map((e) => Be(r, e)), a = t.pkg || Oe("xl/workbook.xml", "", n), o = t.sheetParts && t.sheetParts.length === e.sheets.length ? t.sheetParts : e.sheets.map((e, t) => `xl/worksheets/sheet${t + 1}.xml`);
	i.forEach((e, t) => a.addPart(o[t], e, Ge)), a.addPart("xl/sharedStrings.xml", Ve(r), Ue), a.addPart("xl/styles.xml", He(e), We);
	let s = t.pkg ? qe(t.pkg) : null, c = e.sheets.map((e) => e.name);
	if (s && s.length === c.length && s.every((e, t) => e === c[t])) t.pkg;
	else {
		let t = e.sheets.map((e, t) => `<sheet name="${D(e.name)}" sheetId="${t + 1}" r:id="rId${t + 1}"/>`).join("");
		a.setPart("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<workbook ${M}><bookViews><workbookView activeTab="${e.active}"/></bookViews><sheets>${t}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`);
		let n = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"];
		e.sheets.forEach((e, t) => n.push(`<Relationship Id="rId${t + 1}" Type="${O.rel}/worksheet" Target="worksheets/sheet${t + 1}.xml"/>`));
		let r = e.sheets.length;
		n.push(`<Relationship Id="rId${r + 1}" Type="${O.rel}/styles" Target="styles.xml"/>`), n.push(`<Relationship Id="rId${r + 2}" Type="${O.rel}/sharedStrings" Target="sharedStrings.xml"/>`), n.push("</Relationships>"), a.setPart("xl/_rels/workbook.xml.rels", n.join(""));
	}
	return t.pkg || a.addRel("", O.rel + "/officeDocument", "xl/workbook.xml"), a.save(n.mime);
}
function qe(e) {
	if (!e.has("xl/workbook.xml")) return null;
	let t = e.xml("xl/workbook.xml").children.find((e) => e.localName === "sheets");
	return t ? t.children.filter((e) => e.localName === "sheet").map((e) => e.getAttribute("name") || "") : null;
}
// ../vexcel/src/engine/csv.ts
function Je(e, t = ",") {
	let r = a(), i = r.sheets[0];
	return i.name = "Sheet1", Ye(e, t).forEach((e, t) => {
		e.forEach((e, r) => {
			if (e === "") return;
			let a = Number(e), o = e.trim() !== "" && !isNaN(a) && /^[-+]?[\d.]+(e[-+]?\d+)?$/i.test(e.trim());
			i.cells.set(n(t, r), o ? {
				v: a,
				t: "n"
			} : {
				v: e,
				t: "s"
			});
		});
	}), r;
}
function Ye(e, t) {
	let n = [], r = [], i = "", a = !1;
	for (let o = 0; o < e.length; o++) {
		let s = e[o];
		a ? s === "\"" ? e[o + 1] === "\"" ? (i += "\"", o++) : a = !1 : i += s : s === "\"" ? a = !0 : s === t ? (r.push(i), i = "") : s === "\n" ? (r.push(i), n.push(r), r = [], i = "") : s === "\r" || (i += s);
	}
	return (i !== "" || r.length) && (r.push(i), n.push(r)), n;
}
function Xe(e, t, r = ",") {
	let i = 0, a = 0;
	for (let e of t.cells.keys()) {
		let t = /^([A-Z]+)(\d+)$/.exec(e), n = 0;
		for (let e of t[1]) n = n * 26 + (e.charCodeAt(0) - 64);
		a = Math.max(a, n - 1), i = Math.max(i, t[2] - 1);
	}
	let o = [];
	for (let e = 0; e <= i; e++) {
		let i = [];
		for (let o = 0; o <= a; o++) {
			let a = t.cells.get(n(e, o)), s = a ? String(a.v ?? "") : "";
			(s.includes(r) || s.includes("\"") || s.includes("\n")) && (s = "\"" + s.replace(/"/g, "\"\"") + "\""), i.push(s);
		}
		o.push(i.join(r));
	}
	return o.join("\r\n");
}
// ../vexcel/src/engine/formula.ts
var N = class {
	code;
	constructor(e) {
		this.code = e;
	}
	toString() {
		return this.code;
	}
}, P = () => new N("#DIV/0!"), F = () => new N("#VALUE!"), I = () => new N("#REF!"), Ze = () => new N("#NAME?"), Qe = () => new N("#N/A"), L = () => new N("#NUM!");
function $e(e) {
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
		else throw F();
	}
	return t;
}
function et(e) {
	let n = e.toUpperCase();
	if (n === "TRUE") return {
		k: "bool",
		v: !0
	};
	if (n === "FALSE") return {
		k: "bool",
		v: !1
	};
	if (/^#(DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!|NULL!)$/.test(n)) return {
		k: "err",
		v: n
	};
	let r, i = e, a = e.lastIndexOf("!");
	a >= 0 && (r = e.slice(0, a).replace(/^'(.*)'$/, "$1"), i = e.slice(a + 1));
	let o = i.replace(/\$/g, ""), s = /^([A-Za-z]+)(\d+)$/.exec(o);
	return s ? {
		k: "ref",
		ref: {
			r: s[2] - 1,
			c: t(s[1]),
			sheet: r
		}
	} : null;
}
var tt = {
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
function nt(e) {
	let t = $e(e), n = 0, r = () => t[n], i = () => t[n++];
	function a(e) {
		let t = o();
		for (;;) {
			let n = r();
			if (!n || n.t !== "op" || !(n.v in tt) || tt[n.v] < e) break;
			let o = i().v, s = a(tt[o] + (o === "^" ? 0 : 1));
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
		if (!e) throw F();
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
			if (i()?.t !== "rp") throw F();
			return e;
		}
		if (e.t === "ident") {
			if (r() && r().t === "lp") {
				i();
				let t = [];
				if (r() && r().t !== "rp") for (; t.push(a(1)), r()?.t === "comma";) i();
				if (i()?.t !== "rp") throw F();
				return {
					k: "call",
					name: e.v.toUpperCase(),
					args: t
				};
			}
			if (r() && r().t === "colon") {
				i();
				let t = i(), n = et(e.v), r = t && t.t === "ident" ? et(t.v) : null;
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
				throw I();
			}
			let t = et(e.v);
			if (t) return t;
			throw Ze();
		}
		throw F();
	}
	let l = a(1);
	if (n < t.length) throw F();
	return l;
}
function R(e) {
	if (e instanceof N) return e;
	if (Array.isArray(e)) return e.length ? R(e[0]) : F();
	if (typeof e == "number") return e;
	if (typeof e == "boolean") return +!!e;
	if (e === "") return 0;
	let t = Number(e);
	return isNaN(t) ? F() : t;
}
function z(e) {
	return e instanceof N ? e : Array.isArray(e) ? e.length ? z(e[0]) : F() : typeof e == "boolean" ? e ? "TRUE" : "FALSE" : String(e);
}
function B(e) {
	if (e instanceof N) return e;
	if (Array.isArray(e)) return e.length ? B(e[0]) : F();
	if (typeof e == "boolean") return e;
	if (typeof e == "number") return e !== 0;
	if (e === "TRUE") return !0;
	if (e === "FALSE") return !1;
	let t = Number(e);
	return isNaN(t) ? F() : t !== 0;
}
var V = (e) => e instanceof N;
function H(e) {
	return Array.isArray(e) ? e.flatMap(H) : [e];
}
function U(e) {
	let t = [];
	for (let n of H(e)) typeof n == "number" ? t.push(n) : typeof n == "string" && n !== "" && !isNaN(Number(n)) ? t.push(Number(n)) : typeof n == "boolean" && t.push(+!!n);
	return t;
}
function W(e, t) {
	switch (e.k) {
		case "num": return e.v;
		case "str": return e.v;
		case "bool": return e.v;
		case "err": return new N(e.v);
		case "ref": return t.get(e.ref);
		case "range": return t.range(e.range);
		case "pct": {
			let n = R(W(e.a, t));
			return V(n) ? n : n / 100;
		}
		case "un": {
			let n = R(W(e.a, t));
			return V(n) ? n : e.op === "-" ? -n : n;
		}
		case "bin": {
			let n = W(e.a, t), r = W(e.b, t);
			if (V(n)) return n;
			if (V(r)) return r;
			if (e.op === "&") {
				let e = z(n), t = z(r);
				return V(e) ? e : V(t) ? t : e + t;
			}
			if ([
				"=",
				"<>",
				"<",
				">",
				"<=",
				">="
			].includes(e.op)) return G(n, r, e.op);
			let i = R(n), a = R(r);
			if (V(i)) return i;
			if (V(a)) return a;
			switch (e.op) {
				case "+": return i + a;
				case "-": return i - a;
				case "*": return i * a;
				case "/": return a === 0 ? P() : i / a;
				case "^": {
					let e = i ** +a;
					return isNaN(e) ? L() : e;
				}
			}
			return F();
		}
		case "call": return ft(e.name, e.args.map((e) => W(e, t)), t);
	}
}
function G(e, t, n) {
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
	return F();
}
var K = (e) => {
	for (let t of H(e)) if (V(t)) return t;
	return null;
}, rt = {
	SUM: (e) => K(e) || U(e).reduce((e, t) => e + t, 0),
	AVERAGE: (e) => {
		let t = K(e);
		if (t) return t;
		let n = U(e);
		return n.length ? n.reduce((e, t) => e + t, 0) / n.length : P();
	},
	MIN: (e) => {
		let t = U(e);
		return t.length ? Math.min(...t) : 0;
	},
	MAX: (e) => {
		let t = U(e);
		return t.length ? Math.max(...t) : 0;
	},
	COUNT: (e) => U(e).length,
	COUNTA: (e) => H(e).filter((e) => e !== "" && e != null).length,
	COUNTBLANK: (e) => H(e).filter((e) => e === "" || e == null).length,
	PRODUCT: (e) => {
		let t = U(e);
		return t.length ? t.reduce((e, t) => e * t, 1) : 0;
	},
	SUMPRODUCT: (e) => {
		let t = e.map((e) => H(e)), n = Math.max(...t.map((e) => e.length), 0), r = 0;
		for (let e = 0; e < n; e++) {
			let n = 1;
			for (let r of t) {
				let t = R(r[e] ?? 0);
				if (V(t)) return t;
				n *= t;
			}
			r += n;
		}
		return r;
	},
	ROUND: (e) => {
		let t = R(e[0]), n = R(e[1] ?? 0);
		if (V(t) || V(n)) return F();
		let r = 10 ** n;
		return Math.round(t * r) / r;
	},
	ROUNDUP: (e) => {
		let t = R(e[0]), n = R(e[1] ?? 0);
		if (V(t) || V(n)) return F();
		let r = 10 ** n;
		return (t < 0 ? -1 : 1) * Math.ceil(Math.abs(t) * r) / r;
	},
	ROUNDDOWN: (e) => {
		let t = R(e[0]), n = R(e[1] ?? 0);
		if (V(t) || V(n)) return F();
		let r = 10 ** n;
		return (t < 0 ? -1 : 1) * Math.floor(Math.abs(t) * r) / r;
	},
	INT: (e) => {
		let t = R(e[0]);
		return V(t) ? t : Math.floor(t);
	},
	MOD: (e) => {
		let t = R(e[0]), n = R(e[1]);
		return V(t) || V(n) ? F() : n === 0 ? P() : t - n * Math.floor(t / n);
	},
	ABS: (e) => {
		let t = R(e[0]);
		return V(t) ? t : Math.abs(t);
	},
	SIGN: (e) => {
		let t = R(e[0]);
		return V(t) ? t : Math.sign(t);
	},
	POWER: (e) => {
		let t = R(e[0]), n = R(e[1]);
		if (V(t) || V(n)) return F();
		let r = t ** +n;
		return isNaN(r) ? L() : r;
	},
	SQRT: (e) => {
		let t = R(e[0]);
		return V(t) ? t : t < 0 ? L() : Math.sqrt(t);
	},
	EXP: (e) => {
		let t = R(e[0]);
		return V(t) ? t : Math.exp(t);
	},
	LN: (e) => {
		let t = R(e[0]);
		return V(t) ? t : t <= 0 ? L() : Math.log(t);
	},
	LOG: (e) => {
		let t = R(e[0]), n = R(e[1] ?? 10);
		return V(t) || V(n) ? F() : t <= 0 ? L() : Math.log(t) / Math.log(n);
	},
	LOG10: (e) => {
		let t = R(e[0]);
		return V(t) ? t : t <= 0 ? L() : Math.log10(t);
	},
	SIN: (e) => {
		let t = R(e[0]);
		return V(t) ? t : Math.sin(t);
	},
	COS: (e) => {
		let t = R(e[0]);
		return V(t) ? t : Math.cos(t);
	},
	TAN: (e) => {
		let t = R(e[0]);
		return V(t) ? t : Math.tan(t);
	},
	PI: () => Math.PI,
	RAND: () => Math.random(),
	RANDBETWEEN: (e) => {
		let t = R(e[0]), n = R(e[1]);
		return V(t) || V(n) ? F() : Math.floor(t + Math.random() * (n - t + 1));
	},
	TODAY: () => ut(/* @__PURE__ */ new Date()),
	NOW: () => ut(/* @__PURE__ */ new Date()) + dt(/* @__PURE__ */ new Date()),
	DATE: (e) => {
		let t = R(e[0]), n = R(e[1]), r = R(e[2]);
		return V(t) || V(n) || V(r) ? F() : ut(new Date(t, n - 1, r));
	},
	YEAR: (e) => {
		let t = R(e[0]);
		return V(t) ? t : J(t).getFullYear();
	},
	MONTH: (e) => {
		let t = R(e[0]);
		return V(t) ? t : J(t).getMonth() + 1;
	},
	DAY: (e) => {
		let t = R(e[0]);
		return V(t) ? t : J(t).getDate();
	},
	HOUR: (e) => {
		let t = R(e[0]);
		return V(t) ? t : Math.floor(t % 1 * 24);
	},
	MINUTE: (e) => {
		let t = R(e[0]);
		return V(t) ? t : Math.floor(t % 1 * 24 % 1 * 60);
	},
	SECOND: (e) => {
		let t = R(e[0]);
		return V(t) ? t : Math.round(t % 1 * 24 % 1 * 60 % 1 * 60);
	},
	WEEKDAY: (e) => {
		let t = R(e[0]);
		return V(t) ? t : J(t).getDay() + 1;
	},
	LEFT: (e) => {
		let t = z(e[0]), n = R(e[1] ?? 1);
		return V(t) || V(n) ? F() : t.slice(0, n);
	},
	RIGHT: (e) => {
		let t = z(e[0]), n = R(e[1] ?? 1);
		return V(t) || V(n) ? F() : n === 0 ? "" : t.slice(-n);
	},
	MID: (e) => {
		let t = z(e[0]), n = R(e[1]), r = R(e[2]);
		return V(t) || V(n) || V(r) ? F() : t.substr(n - 1, r);
	},
	LEN: (e) => {
		let t = z(e[0]);
		return V(t) ? t : t.length;
	},
	TRIM: (e) => {
		let t = z(e[0]);
		return V(t) ? t : t.replace(/\s+/g, " ").trim();
	},
	UPPER: (e) => {
		let t = z(e[0]);
		return V(t) ? t : t.toUpperCase();
	},
	LOWER: (e) => {
		let t = z(e[0]);
		return V(t) ? t : t.toLowerCase();
	},
	PROPER: (e) => {
		let t = z(e[0]);
		return V(t) ? t : t.replace(/\w\S*/g, (e) => e[0].toUpperCase() + e.slice(1).toLowerCase());
	},
	SUBSTITUTE: (e) => {
		let t = z(e[0]), n = z(e[1]), r = z(e[2]);
		return V(t) || V(n) || V(r) ? F() : t.split(n).join(r);
	},
	REPLACE: (e) => {
		let t = z(e[0]), n = R(e[1]), r = R(e[2]), i = z(e[3]);
		return V(t) || V(n) || V(r) || V(i) ? F() : t.slice(0, n - 1) + i + t.slice(n - 1 + r);
	},
	FIND: (e) => {
		let t = z(e[0]), n = z(e[1]), r = R(e[2] ?? 1);
		if (V(t) || V(n) || V(r)) return F();
		let i = n.indexOf(t, r - 1);
		return i < 0 ? F() : i + 1;
	},
	SEARCH: (e) => {
		let t = z(e[0]), n = z(e[1]), r = R(e[2] ?? 1);
		if (V(t) || V(n) || V(r)) return F();
		let i = n.toLowerCase().indexOf(t.toLowerCase(), r - 1);
		return i < 0 ? F() : i + 1;
	},
	VALUE: (e) => {
		let t = z(e[0]);
		if (V(t)) return t;
		let n = Number(t);
		return isNaN(n) ? F() : n;
	},
	CONCATENATE: (e) => K(e) || H(e).map((e) => z(e)).filter((e) => !V(e)).join(""),
	CONCAT: (e) => rt.CONCATENATE(e, null),
	EXACT: (e) => {
		let t = z(e[0]), n = z(e[1]);
		return V(t) || V(n) ? F() : t === n;
	},
	REPT: (e) => {
		let t = z(e[0]), n = R(e[1]);
		return V(t) || V(n) ? F() : t.repeat(Math.max(0, Math.floor(n)));
	},
	AND: (e) => K(e) || H(e).every((e) => B(e) === !0),
	OR: (e) => K(e) || H(e).some((e) => B(e) === !0),
	NOT: (e) => {
		let t = B(e[0]);
		return V(t) ? t : !t;
	},
	XOR: (e) => K(e) || H(e).filter((e) => B(e) === !0).length % 2 == 1,
	IF: (e) => {
		let t = B(e[0]);
		return V(t) ? t : t ? e[1] ?? !0 : e[2] ?? !1;
	},
	IFERROR: (e) => V(e[0]) ? e[1] ?? "" : e[0],
	ISBLANK: (e) => e[0] === "" || e[0] == null,
	ISNUMBER: (e) => typeof e[0] == "number",
	ISTEXT: (e) => typeof e[0] == "string",
	ISERROR: (e) => V(e[0]),
	ISNA: (e) => V(e[0]) && e[0].code === "#N/A",
	ISLOGICAL: (e) => typeof e[0] == "boolean",
	CHOOSE: (e) => {
		let t = R(e[0]);
		return V(t) ? t : e[t] ?? F();
	},
	VLOOKUP: (e) => lt(e, q, !0),
	HLOOKUP: (e) => lt(e, q, !1),
	INDEX: (e) => {
		let t = H(e[0]), n = R(e[1] ?? 1), r = R(e[2] ?? 1);
		if (V(n) || V(r)) return F();
		let i = e[0].__dims, a = i ? i[1] : 1;
		return t[(n - 1) * a + (r - 1)] ?? I();
	},
	MATCH: (e) => {
		let t = e[0], n = H(e[1]);
		R(e[2] ?? 1);
		for (let e = 0; e < n.length; e++) if (G(n[e], t, "=") === !0) return e + 1;
		return Qe();
	},
	SUMIF: (e) => ot(e, !1),
	COUNTIF: (e) => at(e),
	AVERAGEIF: (e) => ot(e, !0),
	SUMIFS: (e) => st(e),
	COUNTIFS: (e) => ct(e),
	ROW: (e) => e[0] == null && q.current ? q.current.r + 1 : 0,
	COLUMN: (e) => e[0] == null && q.current ? q.current.c + 1 : 0,
	ROWS: (e) => {
		let t = e[0].__dims;
		return t ? t[0] : Array.isArray(e[0]) ? e[0].length : 1;
	},
	COLUMNS: (e) => {
		let t = e[0].__dims;
		return t ? t[1] : 1;
	},
	OFFSET: (e) => I(),
	INDIRECT: (e) => I(),
	TRANSPOSE: (e) => H(e[0]),
	TEXT: (e) => {
		let t = e[0], n = z(e[1]);
		return V(n) ? n : String(t);
	}
}, q = {
	get: () => I(),
	range: () => []
};
function it(e, t) {
	if (typeof t == "string") {
		let n = /^(<=|>=|<>|<|>|=)(.*)$/.exec(t);
		return n ? G(e, isNaN(Number(n[2])) ? n[2] : Number(n[2]), n[1] === "=" ? "=" : n[1]) === !0 : t.includes("*") || t.includes("?") ? RegExp("^" + t.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i").test(String(e)) : G(e, isNaN(Number(t)) ? t : Number(t), "=") === !0;
	}
	return G(e, t, "=") === !0;
}
function at(e) {
	let t = H(e[0]), n = e[1], r = 0;
	for (let e of t) it(e, n) && r++;
	return r;
}
function ot(e, t) {
	let n = H(e[0]), r = e[1], i = e[2] ? H(e[2]) : n, a = 0, o = 0;
	for (let e = 0; e < n.length; e++) if (it(n[e], r)) {
		let t = R(i[e] ?? 0);
		V(t) || (a += t, o++);
	}
	return t ? o ? a / o : P() : a;
}
function st(e) {
	let t = H(e[0]), n = 0, r = [];
	for (let t = 1; t + 1 < e.length + 1; t += 2) r.push([H(e[t]), e[t + 1]]);
	for (let e = 0; e < t.length; e++) if (r.every(([t, n]) => it(t[e], n))) {
		let r = R(t[e]);
		V(r) || (n += r);
	}
	return n;
}
function ct(e) {
	let t = [];
	for (let n = 0; n + 1 < e.length + 1; n += 2) t.push([H(e[n]), e[n + 1]]);
	let n = t[0]?.[0].length || 0, r = 0;
	for (let e = 0; e < n; e++) t.every(([t, n]) => it(t[e], n)) && r++;
	return r;
}
function lt(e, t, n) {
	let r = e[0], i = e[1], a = R(e[2]), o = e[3] == null || B(e[3]);
	if (V(a)) return F();
	let s = H(i), c = i.__dims, l = c ? c[0] : s.length, u = c ? c[1] : 1;
	if (n) {
		let e = -1;
		for (let t = 0; t < l; t++) {
			let n = s[t * u];
			if (G(n, r, "=") === !0) {
				e = t;
				break;
			}
			o === !0 && G(n, r, "<") === !0 && (e = t);
		}
		return e < 0 ? Qe() : s[e * u + (a - 1)] ?? I();
	}
	{
		let e = -1;
		for (let t = 0; t < u; t++) {
			let n = s[t];
			if (G(n, r, "=") === !0) {
				e = t;
				break;
			}
			o === !0 && G(n, r, "<") === !0 && (e = t);
		}
		return e < 0 ? Qe() : s[(a - 1) * u + e] ?? I();
	}
}
function ut(e) {
	let t = Date.UTC(1899, 11, 30), n = Date.UTC(e.getFullYear(), e.getMonth(), e.getDate()), r = Math.round((n - t) / 864e5);
	return r >= 60 && r++, r;
}
function J(e) {
	let t = Math.floor(e);
	return t >= 60 && t--, new Date(Date.UTC(1899, 11, 30) + t * 864e5);
}
function dt(e) {
	return (e.getHours() * 3600 + e.getMinutes() * 60 + e.getSeconds()) / 86400;
}
function ft(e, t, n) {
	q = n;
	let r = rt[e];
	if (!r) return Ze();
	try {
		return r(t, n);
	} catch (e) {
		return e instanceof N ? e : F();
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
// ../vexcel/src/engine/calc.ts
function X(e, t, n = 0) {
	if (!t) return n;
	let r = e.sheets.findIndex((e) => e.name === t);
	return r >= 0 ? r : n;
}
function pt(e) {
	let t = [], i = /* @__PURE__ */ new Map();
	e.sheets.forEach((a, o) => {
		for (let [s, c] of a.cells) {
			if (!c.f) continue;
			let a = r(s);
			try {
				let r = nt(c.f), l = {
					refs: [],
					ranges: []
				};
				Y(r, l);
				let u = [];
				for (let t of l.refs) u.push(X(e, t.sheet, o) + "!" + n(t.r, t.c));
				for (let t of l.ranges) {
					let r = X(e, t.sheet, o);
					for (let e = t.r1; e <= t.r2; e++) for (let i = t.c1; i <= t.c2; i++) u.push(r + "!" + n(e, i));
				}
				let d = {
					key: o + "!" + s,
					sheet: o,
					r: a.r,
					c: a.c,
					ast: r,
					deps: u
				};
				t.push(d), i.set(d.key, d);
			} catch {
				c.v = "#VALUE!", c.t = "e";
			}
		}
	});
	let a = /* @__PURE__ */ new Map(), o = /* @__PURE__ */ new Map();
	for (let e of t) {
		let t = 0;
		for (let n of e.deps) i.has(n) && (t++, o.set(n, [...o.get(n) || [], e.key]));
		a.set(e.key, t);
	}
	let s = t.filter((e) => (a.get(e.key) || 0) === 0).map((e) => e.key), c = [];
	for (; s.length;) {
		let e = s.shift();
		c.push(e);
		for (let t of o.get(e) || []) {
			let e = (a.get(t) || 0) - 1;
			a.set(t, e), e === 0 && s.push(t);
		}
	}
	let l = new Set(t.filter((e) => !c.includes(e.key)).map((e) => e.key)), u = (t, r, i) => {
		let a = e.sheets[t]?.cells.get(n(r, i));
		return a ? a.v ?? "" : "";
	};
	for (let t of c) {
		let r = i.get(t), a = e.sheets[r.sheet].cells.get(n(r.r, r.c)), o = {
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
		}, s;
		try {
			s = W(r.ast, o);
		} catch (e) {
			s = e instanceof N ? e : new N("#VALUE!");
		}
		Array.isArray(s) && (s = s.length ? s[0] : ""), s instanceof N ? (a.v = s.code, a.t = "e") : typeof s == "number" ? (a.v = s, a.t = "n") : typeof s == "boolean" ? (a.v = s, a.t = "b") : (a.v = typeof s == "string" ? s : String(s), a.t = "s");
	}
	for (let t of l) {
		let r = i.get(t), a = e.sheets[r.sheet].cells.get(n(r.r, r.c));
		a.v = "#REF!", a.t = "e";
	}
}
// ../vexcel/src/engine/numfmt.ts
var mt = [
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
], ht = [
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
], gt = [
	"Sun",
	"Mon",
	"Tue",
	"Wed",
	"Thu",
	"Fri",
	"Sat"
];
function _t(e, t) {
	if (e == null) return "";
	if (typeof e == "boolean") return e ? "TRUE" : "FALSE";
	if (typeof e == "string") return e;
	let n = t || "General";
	return n === "General" ? vt(e) : n === "@" ? String(e) : yt(n) ? bt(e, n) : xt(e, n);
}
function vt(e) {
	if (!isFinite(e) || Number.isInteger(e) && Math.abs(e) < 1e11) return String(e);
	let t = e.toPrecision(11);
	return String(parseFloat(t));
}
function yt(e) {
	return !/[#0]/.test(e) && /[ymdhs]/i.test(e);
}
function Z(e, t) {
	return String(e).padStart(t, "0");
}
function bt(e, t) {
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
			p += t, f.slice(-1), /h$/i.test(m.slice(0, p - t)) || /^s/i.test(m.slice(p)) ? f += t >= 2 ? Z(o, 2) : String(o) : f += t >= 4 ? ht[l] : t === 3 ? mt[l] : t === 2 ? Z(l + 1, 2) : String(l + 1);
		} else if (e === "d" || e === "D") {
			let t = n(e);
			p += t, f += t >= 4 || t === 3 ? gt[d] : t === 2 ? Z(u, 2) : String(u);
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
function xt(e, t) {
	let n = t, r = !1;
	n.includes("%") && (r = !0, e *= 100, n = n.replace(/%/g, "")), n = n.replace(/\[[^\]]*\]/g, ""), n = n.split(";")[0];
	let i = n.indexOf("."), a = i >= 0 ? n.slice(i + 1) : "", o = (a.match(/[0#]/g) || []).length, s = /,/.test(i >= 0 ? n.slice(0, i) : n), c = ((i >= 0 ? n.slice(0, i) : n).match(/^[^0#,]*/) || [""])[0].replace(/"/g, ""), l = (a.match(/[^0#]*$/) || [""])[0].replace(/"/g, "") + (r ? "%" : ""), u = Math.abs(e).toFixed(o);
	if (s) {
		let [e, t] = u.split(".");
		u = e.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (t ? "." + t : "");
	}
	return (e < 0 ? "-" : "") + c + u + l;
}
// ../vexcel/src/engine/grid.ts
var St = 46, Ct = 22;
function wt(e, t = 200, n = 2e3) {
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
function Tt(e, t, n, r) {
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
function Et(e, t, n) {
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
function Dt(t, r) {
	let i = window.devicePixelRatio || 1, a = t.clientWidth, o = t.clientHeight;
	(t.width !== a * i || t.height !== o * i) && (t.width = a * i, t.height = o * i);
	let s = t.getContext("2d");
	s.setTransform(i * r.zoom, 0, 0, i * r.zoom, 0, 0), s.clearRect(0, 0, a / r.zoom, o / r.zoom);
	let { layout: c, sheet: l, wb: u, sel: d } = r, f = 46 - r.scrollX, p = 22 - r.scrollY, m = a / r.zoom, h = o / r.zoom, g = Q("--sheet-grid", "#c0c0c0"), _ = Q("--hdr-a", "#fff"), ee = Q("--hdr-b", "#E4ECF7"), te = Q("--hdr-line", "#9EB6CE"), v = Q("--hdr-ink", "#000"), y = Q("--cell-sel", "rgba(49,74,160,0.2)"), b = Q("--cell-sel-border", "#000"), ne = Math.max(0, $(c.xs, r.scrollX + 46) - 1), re = Math.min(c.nw - 1, $(c.xs, r.scrollX + m) + 1), ie = Math.max(0, $(c.ys, r.scrollY + 22) - 1), ae = Math.min(c.nh - 1, $(c.ys, r.scrollY + h) + 1);
	s.save(), s.beginPath(), s.rect(46, 22, m - 46, h - 22), s.clip(), s.translate(f, p), s.textBaseline = "middle";
	for (let e = ie; e <= ae; e++) for (let t = ne; t <= re; t++) {
		let i = Tt(c, l, e, t), a = l.cells.get(n(e, t)), o = a?.s == null ? void 0 : u.styles[a.s];
		if (o?.fill && (s.fillStyle = o.fill, s.fillRect(i.x, i.y, i.w, i.h)), r.showGrid && (s.strokeStyle = g, s.lineWidth = 1 / r.zoom, s.strokeRect(i.x + .5, i.y + .5, i.w, i.h)), o?.border) {
			s.strokeStyle = "#000", s.lineWidth = 1 / r.zoom;
			let e = o.border;
			s.beginPath(), e.top && (s.moveTo(i.x, i.y + .5), s.lineTo(i.x + i.w, i.y + .5)), e.bottom && (s.moveTo(i.x, i.y + i.h - .5), s.lineTo(i.x + i.w, i.y + i.h - .5)), e.left && (s.moveTo(i.x + .5, i.y), s.lineTo(i.x + .5, i.y + i.h)), e.right && (s.moveTo(i.x + i.w - .5, i.y), s.lineTo(i.x + i.w - .5, i.y + i.h)), s.stroke();
		}
		if (a && a.v != null && a.v !== "") {
			let e = o?.font, t = e?.size || 11;
			s.font = `${e?.italic ? "italic " : ""}${e?.bold ? "bold " : ""}${t}px ${e?.name || "Calibri"}, Carlito, sans-serif`, s.fillStyle = e?.color || Q("--ink", "#000");
			let n = a.t === "e" ? String(a.v) : _t(a.v, o?.numFmt), r = o?.align?.h || (typeof a.v == "number" ? "right" : "left"), c = i.x + 3;
			s.textAlign = "left", r === "right" ? (c = i.x + i.w - 3, s.textAlign = "right") : r === "center" && (c = i.x + i.w / 2, s.textAlign = "center"), s.save(), s.beginPath(), s.rect(i.x, i.y, i.w, i.h), s.clip(), s.fillText(n, c, i.y + i.h / 2), s.restore();
		}
	}
	let x = Ot(d), S = c.xs[x.c1], C = c.ys[x.r1], w = c.xs[x.c2 + 1] - S, oe = c.ys[x.r2 + 1] - C;
	s.fillStyle = y, s.fillRect(S, C, w, oe), s.strokeStyle = b, s.lineWidth = 2 / r.zoom, s.strokeRect(S + 1, C + 1, w - 2, oe - 2);
	let T = Tt(c, l, d.r, d.c);
	s.strokeStyle = b, s.lineWidth = 2 / r.zoom, s.strokeRect(T.x + 1, T.y + 1, T.w - 2, T.h - 2), s.restore(), s.save();
	let E = s.createLinearGradient(0, 0, 0, 22);
	E.addColorStop(0, _), E.addColorStop(1, ee), s.fillStyle = E, s.fillRect(0, 0, m, 22), s.fillRect(0, 0, 46, h), s.strokeStyle = te, s.lineWidth = 1 / r.zoom, s.font = `11px ${Q("--ui-font", "Tahoma")}`, s.fillStyle = v, s.textAlign = "center", s.textBaseline = "middle";
	for (let t = ne; t <= re; t++) {
		let n = c.xs[t] + f, r = c.xs[t + 1] - c.xs[t];
		t >= x.c1 && t <= x.c2 && (s.fillStyle = Q("--hdr-sel-b", "#F8C55E"), s.fillRect(n, 0, r, 22), s.fillStyle = v), s.strokeRect(n + .5, .5, r, 22), s.fillText(e(t), n + r / 2, 11);
	}
	s.textAlign = "center";
	for (let e = ie; e <= ae; e++) {
		let t = c.ys[e] + p, n = c.ys[e + 1] - c.ys[e];
		e >= x.r1 && e <= x.r2 && (s.fillStyle = Q("--hdr-sel-b", "#F8C55E"), s.fillRect(0, t, 46, n), s.fillStyle = v), s.strokeRect(.5, t + .5, 46, n), s.fillText(String(e + 1), 23, t + n / 2);
	}
	s.fillStyle = E, s.fillRect(0, 0, 46, 22), s.strokeRect(.5, .5, 46, 22), s.restore();
}
function $(e, t) {
	let n = 0, r = e.length - 1;
	for (; n < r;) {
		let i = n + r >> 1;
		e[i] < t ? n = i + 1 : r = i;
	}
	return n;
}
function Ot(e) {
	return {
		r1: Math.min(e.r, e.r2),
		c1: Math.min(e.c, e.c2),
		r2: Math.max(e.r, e.r2),
		c2: Math.max(e.c, e.c2)
	};
}
export { c as C, r as S, t as _, Dt as a, i as b, _t as c, Xe as d, Ke as f, u as g, l as h, wt as i, pt as l, n as m, St as n, Ot as o, Fe as p, Tt as r, Et as s, Ct as t, Je as u, e as v, o as w, a as x, s as y };
