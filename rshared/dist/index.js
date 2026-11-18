import { Component as e, createElement as t } from "react";
// ../voffice-shared/src/core/zip.ts
var n = new TextEncoder(), r = new TextDecoder("utf-8"), i = (() => {
	let e = /* @__PURE__ */ new Uint32Array(256);
	for (let t = 0; t < 256; t++) {
		let n = t;
		for (let e = 0; e < 8; e++) n = n & 1 ? 3988292384 ^ n >>> 1 : n >>> 1;
		e[t] = n >>> 0;
	}
	return e;
})();
function a(e) {
	let t = 4294967295;
	for (let n = 0; n < e.length; n++) t = i[(t ^ e[n]) & 255] ^ t >>> 8;
	return (t ^ 4294967295) >>> 0;
}
var o = [
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
], s = [
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
], c = [
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
], l = [
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
], u = [
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
function d(e, t) {
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
var f = null, p = null;
function m() {
	if (f) return;
	let e = /* @__PURE__ */ new Uint8Array(288);
	for (let t = 0; t < 144; t++) e[t] = 8;
	for (let t = 144; t < 256; t++) e[t] = 9;
	for (let t = 256; t < 280; t++) e[t] = 7;
	for (let t = 280; t < 288; t++) e[t] = 8;
	f = d(e, 288), p = d((/* @__PURE__ */ new Uint8Array(30)).fill(5), 30);
}
function h(e, t) {
	let n = 0, r = 0, i = 0, a = new Uint8Array(Math.max(1024, t || e.length * 4)), h = 0, g = (e) => {
		if (h + e <= a.length) return;
		let t = a.length * 2;
		for (; t < h + e;) t *= 2;
		let n = new Uint8Array(t);
		n.set(a.subarray(0, h)), a = n;
	}, _ = (t) => {
		for (; i < t;) {
			if (n >= e.length) throw Error("inflate: unexpected end of data");
			r |= e[n++] << i, i += 8;
		}
		let a = r & (1 << t) - 1;
		return r >>>= t, i -= t, a;
	}, v = (e) => {
		let t = 0, n = 0, r = 0;
		for (let i = 1; i < 16; i++) {
			t |= _(1);
			let a = e.counts[i];
			if (t - a < n) return e.syms[r + (t - n)];
			r += a, n += a, n <<= 1, t <<= 1;
		}
		throw Error("inflate: bad code");
	}, y = (e, t) => {
		for (;;) {
			let n = v(e);
			if (n < 256) g(1), a[h++] = n;
			else if (n === 256) return;
			else {
				if (n -= 257, n >= 29) throw Error("inflate: bad length");
				let e = o[n] + _(s[n]), r = v(t), i = c[r] + _(l[r]);
				if (i > h) throw Error("inflate: distance too far");
				g(e);
				for (let t = 0; t < e; t++) a[h] = a[h - i], h++;
			}
		}
	}, b;
	do {
		b = _(1);
		let t = _(2);
		if (t === 0) {
			r = 0, i = 0;
			let t = e[n] | e[n + 1] << 8;
			n += 4, g(t), a.set(e.subarray(n, n + t), h), h += t, n += t;
		} else if (t === 1) m(), y(f, p);
		else if (t === 2) {
			let e = _(5) + 257, t = _(5) + 1, n = _(4) + 4, r = /* @__PURE__ */ new Uint8Array(19);
			for (let e = 0; e < n; e++) r[u[e]] = _(3);
			let i = d(r, 19), a = new Uint8Array(e + t), o = 0;
			for (; o < e + t;) {
				let e = v(i);
				if (e < 16) a[o++] = e;
				else {
					let t = 0, n = 0;
					if (e === 16) {
						if (!o) throw Error("inflate: repeat with no prior");
						n = a[o - 1], t = 3 + _(2);
					} else t = e === 17 ? 3 + _(3) : 11 + _(7);
					for (; t--;) a[o++] = n;
				}
			}
			y(d(a.subarray(0, e), e), d(a.subarray(e), t));
		} else throw Error("inflate: bad block type");
	} while (!b);
	return a.slice(0, h);
}
async function g(e, t) {
	let n = new Response(new Blob([e]).stream().pipeThrough(t));
	return new Uint8Array(await n.arrayBuffer());
}
async function _(e, t) {
	if (typeof DecompressionStream == "function") try {
		return await g(e, new DecompressionStream("deflate-raw"));
	} catch {}
	return h(e, t);
}
async function v(e) {
	if (typeof CompressionStream == "function") try {
		return await g(e, new CompressionStream("deflate-raw"));
	} catch {}
	return null;
}
function y(e, t, n, i, a, o, s, c, l) {
	let u = null;
	return {
		name: n,
		method: i,
		csize: a,
		usize: o,
		loff: s,
		crc: l,
		async bytes() {
			if (u) return u;
			let r = c;
			if (r == null) {
				let i = s;
				if (i + 30 > e.length || t.getUint32(i, !0) !== 67324752) throw Error("Damaged entry " + n);
				r = i + 30 + t.getUint16(i + 26, !0) + t.getUint16(i + 28, !0);
			}
			let l = e.subarray(r, Math.min(e.length, r + a)), d;
			if (i === 0) d = l.slice();
			else if (i === 8) d = await _(l, o);
			else throw Error("Unsupported compression method " + i + " in " + n);
			return u = d, d;
		},
		async text() {
			return r.decode(await this.bytes());
		}
	};
}
var b = (e, t) => t & 2048 ? r.decode(e) : Array.from(e, (e) => String.fromCharCode(e)).join("");
function x(e, t) {
	if (t.name.endsWith("/")) return;
	e.set(t.name, t);
	let n = t.name.replace(/\\/g, "/");
	n !== t.name && !e.has(n) && e.set(n, t);
}
function ee(e, t) {
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
		let n = t.getUint16(o + 8, !0), r = t.getUint16(o + 10, !0), i = t.getUint32(o + 16, !0), s = t.getUint32(o + 20, !0), c = t.getUint32(o + 24, !0), l = t.getUint16(o + 28, !0), u = t.getUint16(o + 30, !0), d = t.getUint16(o + 32, !0), f = t.getUint32(o + 42, !0), p = b(e.subarray(o + 46, o + 46 + l), n), m = o + 46 + l, h = Math.min(e.length, m + u);
		for (; m + 4 <= h;) {
			let e = t.getUint16(m, !0), n = t.getUint16(m + 2, !0);
			if (e === 1) {
				let e = m + 4;
				c === 4294967295 && (c = Number(t.getBigUint64(e, !0)), e += 8), s === 4294967295 && (s = Number(t.getBigUint64(e, !0)), e += 8), f === 4294967295 && (f = Number(t.getBigUint64(e, !0)));
			}
			m += 4 + n;
		}
		o += 46 + l + u + d, x(a, y(e, t, p, r, s, c, f, void 0, i));
	}
	return a.size ? a : null;
}
function S(e, t) {
	let n = /* @__PURE__ */ new Map(), r = 0;
	for (; r + 30 <= e.length;) {
		if (t.getUint32(r, !0) !== 67324752) {
			let t = r + 1;
			for (; t + 4 <= e.length && (e[t] !== 80 || e[t + 1] !== 75 || e[t + 2] !== 3 || e[t + 3] !== 4);) t++;
			if (t + 30 > e.length) break;
			r = t;
			continue;
		}
		let i = t.getUint16(r + 6, !0), a = t.getUint16(r + 8, !0), o = t.getUint32(r + 18, !0), s = t.getUint32(r + 22, !0), c = t.getUint16(r + 26, !0), l = t.getUint16(r + 28, !0), u = b(e.subarray(r + 30, r + 30 + c), i), d = r + 30 + c + l, f = d + o;
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
		x(n, y(e, t, u, a, o, s, r, d)), r = Math.max(f, r + 30);
	}
	return n;
}
async function C(e) {
	let t = e instanceof Uint8Array ? e : new Uint8Array(e), n = new DataView(t.buffer, t.byteOffset, t.byteLength);
	if (t.length < 22) throw Error("This file is not a valid Office Open XML (ZIP) package.");
	if (n.getUint32(0, !0) !== 67324752 && n.getUint32(0, !0) === 3759263696) {
		let e = new TextDecoder("utf-16le").decode(t.subarray(0, Math.min(t.length, 1 << 20))), n = /EncryptedPackage/.test(e), r = /* @__PURE__ */ Error(n ? "This file is password-protected." : "This is a binary Office 97–2003 file, not an Office Open XML package.");
		throw r.code = "ole", r.encrypted = n, r.bytes = t, r;
	}
	let r = null;
	try {
		r = ee(t, n);
	} catch {
		r = null;
	}
	if (!r) {
		try {
			r = S(t, n);
		} catch {
			r = null;
		}
		if (!r || !r.size) throw Error(n.getUint32(0, !0) === 67324752 ? "This file is damaged and could not be repaired." : "This file is not a valid Office Open XML (ZIP) package.");
		r.repaired = !0;
	}
	return r;
}
function w(e) {
	return {
		time: e.getHours() << 11 | e.getMinutes() << 5 | Math.floor(e.getSeconds() / 2),
		date: e.getFullYear() - 1980 << 9 | e.getMonth() + 1 << 5 | e.getDate()
	};
}
async function T(e, t) {
	let r = [], i = [], o = 0, { time: s, date: c } = w(/* @__PURE__ */ new Date());
	for (let t of e) {
		let e = n.encode(t.name), l = typeof t.data == "string" ? n.encode(t.data) : t.data, u = a(l), d = null;
		!t.store && l.length > 64 && (d = await v(l));
		let f = d && d.length < l.length ? 8 : 0, p = f === 8 ? d : l, m = new Uint8Array(30 + e.length), h = new DataView(m.buffer);
		h.setUint32(0, 67324752, !0), h.setUint16(4, 20, !0), h.setUint16(6, 2048, !0), h.setUint16(8, f, !0), h.setUint16(10, s, !0), h.setUint16(12, c, !0), h.setUint32(14, u, !0), h.setUint32(18, p.length, !0), h.setUint32(22, l.length, !0), h.setUint16(26, e.length, !0), h.setUint16(28, 0, !0), m.set(e, 30), r.push(m, p);
		let g = new Uint8Array(46 + e.length), _ = new DataView(g.buffer);
		_.setUint32(0, 33639248, !0), _.setUint16(4, 20, !0), _.setUint16(6, 20, !0), _.setUint16(8, 2048, !0), _.setUint16(10, f, !0), _.setUint16(12, s, !0), _.setUint16(14, c, !0), _.setUint32(16, u, !0), _.setUint32(20, p.length, !0), _.setUint32(24, l.length, !0), _.setUint16(28, e.length, !0), _.setUint32(42, o, !0), g.set(e, 46), i.push(g), o += m.length + p.length;
	}
	let l = i.reduce((e, t) => e + t.length, 0), u = /* @__PURE__ */ new Uint8Array(22), d = new DataView(u.buffer);
	return d.setUint32(0, 101010256, !0), d.setUint16(8, e.length, !0), d.setUint16(10, e.length, !0), d.setUint32(12, l, !0), d.setUint32(16, o, !0), new Blob([
		...r,
		...i,
		u
	], { type: t || "application/zip" });
}
var E = (e) => n.encode(e), D = (e) => r.decode(e), O = {
	lt: "<",
	gt: ">",
	amp: "&",
	quot: "\"",
	apos: "'"
};
function k(e) {
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
		return O[t] == null ? e : O[t];
	});
}
var A = (e) => (e = String(e).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"), /[\t\n\r]/.test(e) ? e.replace(/\t/g, "&#9;").replace(/\n/g, "&#10;").replace(/\r/g, "&#13;") : e), j = /* @__PURE__ */ RegExp("[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F\\uFFFE\\uFFFF]", "g"), M = /* @__PURE__ */ RegExp("[\\x00-\\x08\\x0B-\\x1F\\uFFFE\\uFFFF]", "g"), N = (e, t) => String(e).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/_x([0-9A-Fa-f]{4})_/g, "_x005F_x$1_").replace(t ? j : M, (e) => "_x" + e.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0") + "_"), te = (e) => e.indexOf("_x") < 0 ? e : e.replace(/_x([0-9A-Fa-f]{4})_/g, (e, t) => String.fromCharCode(parseInt(t, 16))), P = class {
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
}, F = class e {
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
		this.children = [], this.childNodes = [new P(String(e), this)];
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
		if (this.attrs && (n.attrs = { ...this.attrs }), t) for (let e of this.childNodes) n.appendChild(e.nodeType === 3 ? new P(e.data, n) : e.cloneNode(!0));
		return n;
	}
}, I = /([^\s=\/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
function L(e, t) {
	I.lastIndex = 0;
	let n;
	for (; n = I.exec(e);) {
		let e = t.attrs ??= Object.create(null), r = n[2] === void 0 ? n[3] : n[2];
		e[n[1]] = r.indexOf("&") < 0 ? r : k(r);
	}
}
function R(e) {
	let t = Object.create(null);
	I.lastIndex = 0;
	let n;
	for (; n = I.exec(e);) {
		let e = n[2] === void 0 ? n[3] : n[2];
		t[n[1]] = e.indexOf("&") < 0 ? e : k(e);
	}
	return t;
}
var ne = /* @__PURE__ */ new Set([
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
function z(e, t) {
	e && e.charCodeAt(0) === 65279 && (e = e.slice(1));
	let n = !!t?.keepWhitespace, r = !!t?.strict, i = new F("#document", null), a = i, o = 0, s = e.length, c = (e) => {
		a !== i && (n || !/^\s*$/.test(e) || ne.has(a.localName)) && a.childNodes.push(new P(k(e), a));
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
				a !== i && a.childNodes.push(new P(r, a)), o = n < 0 ? s : n + 3;
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
			let c = e.slice(t + 1, i), l = e.charCodeAt(n - 1) === 47, u = new F(c, a), d = l ? n - 1 : n;
			i < d && L(e.slice(i, d), u), a.childNodes.push(u), a.children.push(u), l || (a = u), o = n + 1;
		}
	}
	let l = i.children[0];
	if (!l) throw Error("XML parse error: no root element");
	if (r && (a !== i || i.children.length !== 1)) throw Error("XML parse error: incomplete document");
	return l;
}
function re(e) {
	return e.indexOf("\r") >= 0 && (e = e.replace(/\r\n?/g, "\n")), e.indexOf("<![CDATA[") < 0 ? k(e) : e.split(/(<!\[CDATA\[[\s\S]*?\]\]>)/).map((e) => e.startsWith("<![CDATA[") ? e.slice(9, -3) : k(e)).join("");
}
function ie(e, t, n, r) {
	let i = /<(\/?)(?:[A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)([^>]*?)(\/?)>/g;
	i.lastIndex = t;
	let a = t, o;
	for (; (o = i.exec(e)) && o.index < n;) {
		let t = o.index > a ? e.slice(a, o.index) : "";
		r(o[2], o[1] ? 2 : o[4] ? 3 : 1, o[3], t), a = i.lastIndex;
	}
}
function B(e) {
	return e.length >= 2 && e[0] === 255 && e[1] === 254 ? new TextDecoder("utf-16le").decode(e.subarray(2)) : e.length >= 2 && e[0] === 254 && e[1] === 255 ? new TextDecoder("utf-16be").decode(e.subarray(2)) : e.length >= 4 && e[0] === 0 && e[1] === 60 && e[2] === 0 && e[3] === 63 ? new TextDecoder("utf-16be").decode(e) : e.length >= 4 && e[0] === 60 && e[1] === 0 && e[2] === 63 && e[3] === 0 ? new TextDecoder("utf-16le").decode(e) : new TextDecoder("utf-8").decode(e);
}
function V(e) {
	if (!e) return "";
	let t = {};
	for (let n = e; n && n.nodeType === 1; n = n.parentNode) if (n.attrs) for (let e in n.attrs) (e === "xmlns" || e.startsWith("xmlns:")) && (t[e] = n.attrs[e]);
	let n = (e, r) => {
		if (e.nodeType === 3) return e.data.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
		let i = e, a = i.attrs ? { ...i.attrs } : {};
		if (r) for (let [e, n] of Object.entries(t)) e in a || (a[e] = n);
		let o = "<" + i.nodeName + Object.entries(a).map(([e, t]) => " " + e + "=\"" + A(t) + "\"").join("");
		return i.childNodes.length ? o + ">" + i.childNodes.map((e) => n(e, !1)).join("") + "</" + i.nodeName + ">" : o + "/>";
	};
	return n(e, !0);
}
function H(e, t, n) {
	let r = new F(e, null);
	if (t) for (let [e, n] of Object.entries(t)) r.setAttribute(e, n);
	return n != null && (r.textContent = n), r;
}
// ../voffice-shared/src/core/opc.ts
var U = Object.freeze({
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
}), W = Object.freeze(Object.fromEntries([
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
}))), G = () => {
	if (crypto.randomUUID) return crypto.randomUUID();
	let e = crypto.getRandomValues(/* @__PURE__ */ new Uint8Array(16));
	e[6] = e[6] & 15 | 64, e[8] = e[8] & 63 | 128;
	let t = Array.from(e, (e) => e.toString(16).padStart(2, "0")).join("");
	return `${t.slice(0, 8)}-${t.slice(8, 12)}-${t.slice(12, 16)}-${t.slice(16, 20)}-${t.slice(20)}`;
}, K = "[Content_Types].xml", q = (e) => {
	let t = e.lastIndexOf("/");
	return (t < 0 ? "" : e.slice(0, t + 1)) + "_rels/" + (t < 0 ? e : e.slice(t + 1)) + ".rels";
};
function J(e, t) {
	if (t.startsWith("/")) return t.slice(1);
	let n = ((e.includes("/") ? e.slice(0, e.lastIndexOf("/") + 1) : "") + t).split("/"), r = [];
	for (let e of n) e !== "." && e !== "" && (e === ".." ? r.pop() : r.push(e));
	return r.join("/");
}
var Y = (e) => e.replace(U.strictRel + "/", U.rel + "/"), X = class e {
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
		let n = await C(t), r = new e();
		r.repaired = !!n.repaired;
		for (let [e, t] of n) r.parts.set(e, await t.bytes());
		let i = r.parts.get(K);
		if (!i) throw Error("This package has no [Content_Types].xml part.");
		let a = z(B(i));
		for (let e of a.children) e.localName === "Default" ? r.defaults.set(e.getAttribute("Extension") || "", e.getAttribute("ContentType") || "") : e.localName === "Override" && r.overrides.set((e.getAttribute("PartName") || "").replace(/^\//, ""), e.getAttribute("ContentType") || "");
		let o = r.relsOf("").find((e) => Y(e.type) === U.rel + "/officeDocument");
		return o && (r.main = J("", o.target)), r;
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
			t = B(n), this.textCache.set(e, t);
		}
		return t;
	}
	xml(e) {
		let t = this.xmlCache.get(e);
		return t || (t = z(this.text(e)), this.xmlCache.set(e, t)), t;
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
		let r = this.parts.get(q(e));
		if (r) {
			let e = z(B(r));
			for (let t of e.children) t.localName === "Relationship" && n.push({
				id: t.getAttribute("Id") || "",
				type: Y(t.getAttribute("Type") || ""),
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
		return this.relsOf(e).filter((e) => e.type === t && e.mode !== "External").map((t) => J(e, t.target));
	}
	setPart(e, t) {
		this.parts.set(e, typeof t == "string" ? new TextEncoder().encode(t) : t), this.edited.add(e), this.textCache.delete(e), this.xmlCache.delete(e);
	}
	addPart(e, t, n) {
		if (this.setPart(e, t), n) this.overrides.set(e, n), this.ctDirty = !0;
		else {
			let t = e.includes(".") ? e.slice(e.lastIndexOf(".") + 1) : "";
			t && !this.defaults.has(t) && (this.defaults.set(t, Z(t)), this.ctDirty = !0);
		}
	}
	removePart(e) {
		this.parts.delete(e), this.edited.delete(e), this.textCache.delete(e), this.xmlCache.delete(e), this.overrides.delete(e) && (this.ctDirty = !0);
		let t = q(e);
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
		return Object.values(W).find((e) => e.contentType === t)?.ext || e;
	}
	writeContentTypes() {
		if (!this.ctDirty) return;
		let e = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", "<Types xmlns=\"" + U.ct + "\">"];
		for (let [t, n] of [...this.defaults].sort()) e.push(`<Default Extension="${A(t)}" ContentType="${A(n)}"/>`);
		for (let [t, n] of [...this.overrides].sort()) e.push(`<Override PartName="/${A(t)}" ContentType="${A(n)}"/>`);
		e.push("</Types>"), this.parts.set(K, new TextEncoder().encode(e.join("")));
	}
	writeRels() {
		for (let e of this.relsDirty) {
			let t = this.relsCache.get(e) || [], n = e, r = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", "<Relationships xmlns=\"" + U.pkg + "\">"];
			for (let e of t) r.push(`<Relationship Id="${A(e.id)}" Type="${A(e.type)}" Target="${A(e.target)}"${e.mode ? ` TargetMode="${A(e.mode)}"` : ""}/>`);
			r.push("</Relationships>");
			let i = q(n);
			this.parts.set(i, new TextEncoder().encode(r.join(""))), this.overrides.has(i) || (this.defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml"), this.ctDirty = !0);
		}
	}
	async save(e) {
		return this.writeRels(), this.writeContentTypes(), T([...this.parts].map(([e, t]) => ({
			name: e,
			data: t
		})), e || "application/zip");
	}
};
function Z(e) {
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
function ae(e, t, n, r = {}) {
	let i = new X();
	i.main = e, i.defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml"), i.defaults.set("xml", "application/xml");
	for (let [e, t] of Object.entries(r)) i.defaults.set(e, t);
	return i.overrides.set(e, n.contentType), i.setPart(e, t), i.addRel("", U.rel + "/officeDocument", e), i;
}
function oe(e) {
	return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" + V(e);
}
// src/index.ts
function Q(e) {
	try {
		return decodeURIComponent(new URL(e).pathname.split("/").filter(Boolean).pop() || "document");
	} catch {
		return "document";
	}
}
function $(e) {
	return e.token ? { Authorization: `Bearer ${e.token}` } : e.username == null ? {} : { Authorization: `Basic ${btoa(unescape(encodeURIComponent(`${e.username}:${e.password ?? ""}`)))}` };
}
async function se(e) {
	let t = e.kind === "webdav" ? $(e) : e.kind === "http" ? e.headers : void 0, n = e.kind === "s3" ? `${e.endpoint.replace(/\/$/, "")}/${e.forcePathStyle ? `${e.bucket}/` : ""}${e.key}` : e.url, r = await fetch(n, { headers: t });
	if (!r.ok) throw Error(`Download failed: ${r.status} ${r.statusText}`);
	return {
		blob: await r.blob(),
		name: Q(n)
	};
}
async function ce(e, t, n) {
	if (t = await t, e.kind === "post") {
		let r = new FormData();
		Object.entries(e.fields ?? {}).forEach(([e, t]) => r.append(e, t)), r.append(e.field ?? "file", t, n);
		let i = await fetch(e.url, {
			method: "POST",
			headers: e.headers,
			body: r
		});
		if (!i.ok) throw Error(`Upload failed: ${i.status} ${i.statusText}`);
		return;
	}
	let r = e.kind === "s3" ? `${e.endpoint.replace(/\/$/, "")}/${e.forcePathStyle ? `${e.bucket}/` : ""}${e.key}` : e.kind === "presigned" ? e.putUrl : e.url;
	if (!r) throw Error("The online source is read-only");
	let i = await fetch(r, {
		method: (e.kind, "PUT"),
		headers: e.kind === "webdav" ? $(e) : void 0,
		body: t
	});
	if (!i.ok) throw Error(`Upload failed: ${i.status} ${i.statusText}`);
}
function le(e, t) {
	let n = document.createElement("a");
	n.href = URL.createObjectURL(e), n.download = t, n.click(), setTimeout(() => URL.revokeObjectURL(n.href), 0);
}
var ue = class extends e {
	state = { error: null };
	static getDerivedStateFromError(e) {
		return { error: e };
	}
	componentDidCatch(e, t) {
		console.error(e, t);
	}
	render() {
		return this.state.error ? t("div", {
			role: "alert",
			className: "vo-error"
		}, this.state.error.message) : this.props.children;
	}
};
export { F as El, ue as ErrorBoundary, U as NS, X as Package, P as Txt, R as attrs, re as chars, Z as contentTypeForExt, a as crc32, D as decode, B as decodeBytes, le as downloadBlob, E as encode, A as esc, N as escText, se as fetchOnline, W as formats, _ as inflateRaw, h as inflateRawJS, H as mkEl, Q as nameFromUrl, ae as newPackage, z as parseXml, C as readZip, J as resolveTarget, ce as saveOnline, ie as scanXml, V as serializeXml, k as unescape, te as unx, G as uuid, T as writeZip, oe as xmlPart };
