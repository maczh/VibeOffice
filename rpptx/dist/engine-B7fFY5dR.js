// ../vpptx/src/engine/model.ts
var e = 0, t = (t = "s") => `${t}${++e}_${Math.random().toString(36).slice(2, 7)}`, n = 9525, r = (e) => e / n, i = (e) => Math.round(e * n), a = () => ({
	id: t("sl"),
	shapes: []
}), o = () => ({
	slideW: 12192e3,
	slideH: 6858e3,
	slides: [a()],
	format: "pptx"
});
function s(e, n, r, i, a = "") {
	return {
		id: t("sh"),
		kind: "sp",
		x: e,
		y: n,
		w: r,
		h: i,
		preset: "rect",
		fill: null,
		line: null,
		text: [{
			align: "left",
			runs: a ? [{ text: a }] : []
		}]
	};
}
var c = (e) => JSON.parse(JSON.stringify(e)), l = new TextEncoder(), u = new TextDecoder("utf-8"), d = (() => {
	let e = /* @__PURE__ */ new Uint32Array(256);
	for (let t = 0; t < 256; t++) {
		let n = t;
		for (let e = 0; e < 8; e++) n = n & 1 ? 3988292384 ^ n >>> 1 : n >>> 1;
		e[t] = n >>> 0;
	}
	return e;
})();
function f(e) {
	let t = 4294967295;
	for (let n = 0; n < e.length; n++) t = d[(t ^ e[n]) & 255] ^ t >>> 8;
	return (t ^ 4294967295) >>> 0;
}
var p = [
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
], m = [
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
], h = [
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
], g = [
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
], _ = [
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
var y = null, ee = null;
function te() {
	if (y) return;
	let e = /* @__PURE__ */ new Uint8Array(288);
	for (let t = 0; t < 144; t++) e[t] = 8;
	for (let t = 144; t < 256; t++) e[t] = 9;
	for (let t = 256; t < 280; t++) e[t] = 7;
	for (let t = 280; t < 288; t++) e[t] = 8;
	y = v(e, 288), ee = v((/* @__PURE__ */ new Uint8Array(30)).fill(5), 30);
}
function ne(e, t) {
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
				let e = p[n] + c(m[n]), r = l(t), i = h[r] + c(g[r]);
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
		} else if (t === 1) te(), u(y, ee);
		else if (t === 2) {
			let e = c(5) + 257, t = c(5) + 1, n = c(4) + 4, r = /* @__PURE__ */ new Uint8Array(19);
			for (let e = 0; e < n; e++) r[_[e]] = c(3);
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
async function b(e, t) {
	let n = new Response(new Blob([e]).stream().pipeThrough(t));
	return new Uint8Array(await n.arrayBuffer());
}
async function re(e, t) {
	if (typeof DecompressionStream == "function") try {
		return await b(e, new DecompressionStream("deflate-raw"));
	} catch {}
	return ne(e, t);
}
async function ie(e) {
	if (typeof CompressionStream == "function") try {
		return await b(e, new CompressionStream("deflate-raw"));
	} catch {}
	return null;
}
function x(e, t, n, r, i, a, o, s, c) {
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
			else if (r === 8) d = await re(u, a);
			else throw Error("Unsupported compression method " + r + " in " + n);
			return l = d, d;
		},
		async text() {
			return u.decode(await this.bytes());
		}
	};
}
var S = (e, t) => t & 2048 ? u.decode(e) : Array.from(e, (e) => String.fromCharCode(e)).join("");
function C(e, t) {
	if (t.name.endsWith("/")) return;
	e.set(t.name, t);
	let n = t.name.replace(/\\/g, "/");
	n !== t.name && !e.has(n) && e.set(n, t);
}
function ae(e, t) {
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
		let n = t.getUint16(o + 8, !0), r = t.getUint16(o + 10, !0), i = t.getUint32(o + 16, !0), s = t.getUint32(o + 20, !0), c = t.getUint32(o + 24, !0), l = t.getUint16(o + 28, !0), u = t.getUint16(o + 30, !0), d = t.getUint16(o + 32, !0), f = t.getUint32(o + 42, !0), p = S(e.subarray(o + 46, o + 46 + l), n), m = o + 46 + l, h = Math.min(e.length, m + u);
		for (; m + 4 <= h;) {
			let e = t.getUint16(m, !0), n = t.getUint16(m + 2, !0);
			if (e === 1) {
				let e = m + 4;
				c === 4294967295 && (c = Number(t.getBigUint64(e, !0)), e += 8), s === 4294967295 && (s = Number(t.getBigUint64(e, !0)), e += 8), f === 4294967295 && (f = Number(t.getBigUint64(e, !0)));
			}
			m += 4 + n;
		}
		o += 46 + l + u + d, C(a, x(e, t, p, r, s, c, f, void 0, i));
	}
	return a.size ? a : null;
}
function oe(e, t) {
	let n = /* @__PURE__ */ new Map(), r = 0;
	for (; r + 30 <= e.length;) {
		if (t.getUint32(r, !0) !== 67324752) {
			let t = r + 1;
			for (; t + 4 <= e.length && (e[t] !== 80 || e[t + 1] !== 75 || e[t + 2] !== 3 || e[t + 3] !== 4);) t++;
			if (t + 30 > e.length) break;
			r = t;
			continue;
		}
		let i = t.getUint16(r + 6, !0), a = t.getUint16(r + 8, !0), o = t.getUint32(r + 18, !0), s = t.getUint32(r + 22, !0), c = t.getUint16(r + 26, !0), l = t.getUint16(r + 28, !0), u = S(e.subarray(r + 30, r + 30 + c), i), d = r + 30 + c + l, f = d + o;
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
		C(n, x(e, t, u, a, o, s, r, d)), r = Math.max(f, r + 30);
	}
	return n;
}
async function se(e) {
	let t = e instanceof Uint8Array ? e : new Uint8Array(e), n = new DataView(t.buffer, t.byteOffset, t.byteLength);
	if (t.length < 22) throw Error("This file is not a valid Office Open XML (ZIP) package.");
	if (n.getUint32(0, !0) !== 67324752 && n.getUint32(0, !0) === 3759263696) {
		let e = new TextDecoder("utf-16le").decode(t.subarray(0, Math.min(t.length, 1 << 20))), n = /EncryptedPackage/.test(e), r = /* @__PURE__ */ Error(n ? "This file is password-protected." : "This is a binary Office 97–2003 file, not an Office Open XML package.");
		throw r.code = "ole", r.encrypted = n, r.bytes = t, r;
	}
	let r = null;
	try {
		r = ae(t, n);
	} catch {
		r = null;
	}
	if (!r) {
		try {
			r = oe(t, n);
		} catch {
			r = null;
		}
		if (!r || !r.size) throw Error(n.getUint32(0, !0) === 67324752 ? "This file is damaged and could not be repaired." : "This file is not a valid Office Open XML (ZIP) package.");
		r.repaired = !0;
	}
	return r;
}
function ce(e) {
	return {
		time: e.getHours() << 11 | e.getMinutes() << 5 | Math.floor(e.getSeconds() / 2),
		date: e.getFullYear() - 1980 << 9 | e.getMonth() + 1 << 5 | e.getDate()
	};
}
async function le(e, t) {
	let n = [], r = [], i = 0, { time: a, date: o } = ce(/* @__PURE__ */ new Date());
	for (let t of e) {
		let e = l.encode(t.name), s = typeof t.data == "string" ? l.encode(t.data) : t.data, c = f(s), u = null;
		!t.store && s.length > 64 && (u = await ie(s));
		let d = u && u.length < s.length ? 8 : 0, p = d === 8 ? u : s, m = new Uint8Array(30 + e.length), h = new DataView(m.buffer);
		h.setUint32(0, 67324752, !0), h.setUint16(4, 20, !0), h.setUint16(6, 2048, !0), h.setUint16(8, d, !0), h.setUint16(10, a, !0), h.setUint16(12, o, !0), h.setUint32(14, c, !0), h.setUint32(18, p.length, !0), h.setUint32(22, s.length, !0), h.setUint16(26, e.length, !0), h.setUint16(28, 0, !0), m.set(e, 30), n.push(m, p);
		let g = new Uint8Array(46 + e.length), _ = new DataView(g.buffer);
		_.setUint32(0, 33639248, !0), _.setUint16(4, 20, !0), _.setUint16(6, 20, !0), _.setUint16(8, 2048, !0), _.setUint16(10, d, !0), _.setUint16(12, a, !0), _.setUint16(14, o, !0), _.setUint32(16, c, !0), _.setUint32(20, p.length, !0), _.setUint32(24, s.length, !0), _.setUint16(28, e.length, !0), _.setUint32(42, i, !0), g.set(e, 46), r.push(g), i += m.length + p.length;
	}
	let s = r.reduce((e, t) => e + t.length, 0), c = /* @__PURE__ */ new Uint8Array(22), u = new DataView(c.buffer);
	return u.setUint32(0, 101010256, !0), u.setUint16(8, e.length, !0), u.setUint16(10, e.length, !0), u.setUint32(12, s, !0), u.setUint32(16, i, !0), new Blob([
		...n,
		...r,
		c
	], { type: t || "application/zip" });
}
// ../voffice-shared/src/core/xml.ts
var w = {
	lt: "<",
	gt: ">",
	amp: "&",
	quot: "\"",
	apos: "'"
};
function T(e) {
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
		return w[t] == null ? e : w[t];
	});
}
var E = (e) => (e = String(e).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"), /[\t\n\r]/.test(e) ? e.replace(/\t/g, "&#9;").replace(/\n/g, "&#10;").replace(/\r/g, "&#13;") : e), ue = /* @__PURE__ */ RegExp("[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F\\uFFFE\\uFFFF]", "g"), de = /* @__PURE__ */ RegExp("[\\x00-\\x08\\x0B-\\x1F\\uFFFE\\uFFFF]", "g"), fe = (e, t) => String(e).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/_x([0-9A-Fa-f]{4})_/g, "_x005F_x$1_").replace(t ? ue : de, (e) => "_x" + e.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0") + "_"), D = class {
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
}, O = class e {
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
		this.children = [], this.childNodes = [new D(String(e), this)];
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
		if (this.attrs && (n.attrs = { ...this.attrs }), t) for (let e of this.childNodes) n.appendChild(e.nodeType === 3 ? new D(e.data, n) : e.cloneNode(!0));
		return n;
	}
}, k = /([^\s=\/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
function pe(e, t) {
	k.lastIndex = 0;
	let n;
	for (; n = k.exec(e);) {
		let e = t.attrs ??= Object.create(null), r = n[2] === void 0 ? n[3] : n[2];
		e[n[1]] = r.indexOf("&") < 0 ? r : T(r);
	}
}
var me = /* @__PURE__ */ new Set([
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
function A(e, t) {
	e && e.charCodeAt(0) === 65279 && (e = e.slice(1));
	let n = !!t?.keepWhitespace, r = !!t?.strict, i = new O("#document", null), a = i, o = 0, s = e.length, c = (e) => {
		a !== i && (n || !/^\s*$/.test(e) || me.has(a.localName)) && a.childNodes.push(new D(T(e), a));
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
				a !== i && a.childNodes.push(new D(r, a)), o = n < 0 ? s : n + 3;
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
			let c = e.slice(t + 1, i), l = e.charCodeAt(n - 1) === 47, u = new O(c, a), d = l ? n - 1 : n;
			i < d && pe(e.slice(i, d), u), a.childNodes.push(u), a.children.push(u), l || (a = u), o = n + 1;
		}
	}
	let l = i.children[0];
	if (!l) throw Error("XML parse error: no root element");
	if (r && (a !== i || i.children.length !== 1)) throw Error("XML parse error: incomplete document");
	return l;
}
function j(e) {
	return e.length >= 2 && e[0] === 255 && e[1] === 254 ? new TextDecoder("utf-16le").decode(e.subarray(2)) : e.length >= 2 && e[0] === 254 && e[1] === 255 ? new TextDecoder("utf-16be").decode(e.subarray(2)) : e.length >= 4 && e[0] === 0 && e[1] === 60 && e[2] === 0 && e[3] === 63 ? new TextDecoder("utf-16be").decode(e) : e.length >= 4 && e[0] === 60 && e[1] === 0 && e[2] === 63 && e[3] === 0 ? new TextDecoder("utf-16le").decode(e) : new TextDecoder("utf-8").decode(e);
}
// ../voffice-shared/src/core/opc.ts
var M = Object.freeze({
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
}), N = Object.freeze(Object.fromEntries([
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
}))), P = "[Content_Types].xml", F = (e) => {
	let t = e.lastIndexOf("/");
	return (t < 0 ? "" : e.slice(0, t + 1)) + "_rels/" + (t < 0 ? e : e.slice(t + 1)) + ".rels";
};
function I(e, t) {
	if (t.startsWith("/")) return t.slice(1);
	let n = ((e.includes("/") ? e.slice(0, e.lastIndexOf("/") + 1) : "") + t).split("/"), r = [];
	for (let e of n) e !== "." && e !== "" && (e === ".." ? r.pop() : r.push(e));
	return r.join("/");
}
var L = (e) => e.replace(M.strictRel + "/", M.rel + "/"), R = class e {
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
		let n = await se(t), r = new e();
		r.repaired = !!n.repaired;
		for (let [e, t] of n) r.parts.set(e, await t.bytes());
		let i = r.parts.get(P);
		if (!i) throw Error("This package has no [Content_Types].xml part.");
		let a = A(j(i));
		for (let e of a.children) e.localName === "Default" ? r.defaults.set(e.getAttribute("Extension") || "", e.getAttribute("ContentType") || "") : e.localName === "Override" && r.overrides.set((e.getAttribute("PartName") || "").replace(/^\//, ""), e.getAttribute("ContentType") || "");
		let o = r.relsOf("").find((e) => L(e.type) === M.rel + "/officeDocument");
		return o && (r.main = I("", o.target)), r;
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
			t = j(n), this.textCache.set(e, t);
		}
		return t;
	}
	xml(e) {
		let t = this.xmlCache.get(e);
		return t || (t = A(this.text(e)), this.xmlCache.set(e, t)), t;
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
		let r = this.parts.get(F(e));
		if (r) {
			let e = A(j(r));
			for (let t of e.children) t.localName === "Relationship" && n.push({
				id: t.getAttribute("Id") || "",
				type: L(t.getAttribute("Type") || ""),
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
		return this.relsOf(e).filter((e) => e.type === t && e.mode !== "External").map((t) => I(e, t.target));
	}
	setPart(e, t) {
		this.parts.set(e, typeof t == "string" ? new TextEncoder().encode(t) : t), this.edited.add(e), this.textCache.delete(e), this.xmlCache.delete(e);
	}
	addPart(e, t, n) {
		if (this.setPart(e, t), n) this.overrides.set(e, n), this.ctDirty = !0;
		else {
			let t = e.includes(".") ? e.slice(e.lastIndexOf(".") + 1) : "";
			t && !this.defaults.has(t) && (this.defaults.set(t, he(t)), this.ctDirty = !0);
		}
	}
	removePart(e) {
		this.parts.delete(e), this.edited.delete(e), this.textCache.delete(e), this.xmlCache.delete(e), this.overrides.delete(e) && (this.ctDirty = !0);
		let t = F(e);
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
		return Object.values(N).find((e) => e.contentType === t)?.ext || e;
	}
	writeContentTypes() {
		if (!this.ctDirty) return;
		let e = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", "<Types xmlns=\"" + M.ct + "\">"];
		for (let [t, n] of [...this.defaults].sort()) e.push(`<Default Extension="${E(t)}" ContentType="${E(n)}"/>`);
		for (let [t, n] of [...this.overrides].sort()) e.push(`<Override PartName="/${E(t)}" ContentType="${E(n)}"/>`);
		e.push("</Types>"), this.parts.set(P, new TextEncoder().encode(e.join("")));
	}
	writeRels() {
		for (let e of this.relsDirty) {
			let t = this.relsCache.get(e) || [], n = e, r = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", "<Relationships xmlns=\"" + M.pkg + "\">"];
			for (let e of t) r.push(`<Relationship Id="${E(e.id)}" Type="${E(e.type)}" Target="${E(e.target)}"${e.mode ? ` TargetMode="${E(e.mode)}"` : ""}/>`);
			r.push("</Relationships>");
			let i = F(n);
			this.parts.set(i, new TextEncoder().encode(r.join(""))), this.overrides.has(i) || (this.defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml"), this.ctDirty = !0);
		}
	}
	async save(e) {
		return this.writeRels(), this.writeContentTypes(), le([...this.parts].map(([e, t]) => ({
			name: e,
			data: t
		})), e || "application/zip");
	}
};
function he(e) {
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
function ge(e, t, n, r = {}) {
	let i = new R();
	i.main = e, i.defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml"), i.defaults.set("xml", "application/xml");
	for (let [e, t] of Object.entries(r)) i.defaults.set(e, t);
	return i.overrides.set(e, n.contentType), i.setPart(e, t), i.addRel("", M.rel + "/officeDocument", e), i;
}
M.p, M.a;
var z = M.r, B = (e, t) => e.children.filter((e) => e.localName === t), V = (e, t) => e.children.find((e) => e.localName === t), H = (e, t) => e ? e.getAttribute(t) : null, U = (e, t) => e.getElementsByTagNameNS("*", t)[0];
function W(e) {
	let t = e.attributes.map((e) => ` ${e.name}="${e.value.replace(/"/g, "&quot;")}"`).join(""), n = e.childNodes.map((e) => e.nodeType === 3 ? e.data.replace(/&/g, "&amp;").replace(/</g, "&lt;") : W(e)).join("");
	return `<${e.nodeName}${t}>${n}</${e.nodeName}>`;
}
function G(e) {
	if (!e) return;
	let t = {}, n = H(e, "sz");
	n && (t.size = n / 100), H(e, "b") === "1" && (t.bold = !0), H(e, "i") === "1" && (t.italic = !0), H(e, "u") && H(e, "u") !== "none" && (t.underline = !0);
	let r = V(e, "solidFill"), i = r && (V(r, "srgbClr") || V(r, "schemeClr"));
	i && V(r, "srgbClr") && (t.color = ("#" + H(i, "val")).toLowerCase());
	let a = V(e, "latin");
	return a && (t.font = H(a, "typeface") || void 0), Object.keys(t).length ? t : void 0;
}
function _e(e) {
	if (!e) return [];
	let t = [];
	for (let n of B(e, "p")) {
		let e = { runs: [] }, r = V(n, "pPr"), i = H(r, "algn");
		i && (e.align = i === "ctr" ? "center" : i === "r" ? "right" : "left"), r && (V(r, "buChar") || V(r, "buAutoNum")) && (e.bullet = !0);
		for (let t of B(n, "r")) {
			let n = V(t, "t");
			e.runs.push({
				text: n?.textContent || "",
				rPr: G(V(t, "rPr"))
			});
		}
		let a = V(n, "endParaRPr");
		!e.runs.length && a && e.runs.push({
			text: "",
			rPr: G(a)
		}), t.push(e);
	}
	return t;
}
function ve(e) {
	let t = e && V(e, "xfrm"), n = t && V(t, "off"), r = t && V(t, "ext");
	return {
		x: +(H(n, "x") || 0),
		y: +(H(n, "y") || 0),
		w: +(H(r, "cx") || 0),
		h: +(H(r, "cy") || 0),
		rot: t && H(t, "rot") ? H(t, "rot") / 6e4 : void 0
	};
}
function ye(e) {
	let t, n;
	if (!e) return {
		fill: t,
		line: n
	};
	V(e, "noFill") && (t = null);
	let r = V(e, "solidFill");
	if (r) {
		let e = V(r, "srgbClr");
		e && (t = { color: ("#" + H(e, "val")).toLowerCase() });
	}
	let i = V(e, "ln");
	if (i) {
		if (V(i, "noFill")) n = null;
		else {
			let e = V(i, "solidFill"), t = e && V(e, "srgbClr");
			n = {
				color: t ? ("#" + H(t, "val")).toLowerCase() : "#000000",
				width: (H(i, "w") || 12700) / 12700
			};
			let r = V(i, "prstDash");
			r && (n.dash = H(r, "val") || void 0);
		}
	}
	return {
		fill: t,
		line: n
	};
}
function be(e, n, r) {
	let i = e.localName, a = V(e, "nvSpPr") || V(e, "nvPicPr") || V(e, "nvCxnSpPr") || V(e, "nvGraphicFramePr"), o = H(a && V(a, "cNvPr"), "name") || "", s = V(e, "spPr"), c = ve(s);
	if (i === "pic") {
		let i = U(e, "blip"), a = i && (i.getAttributeNS(z, "embed") || H(i, "r:embed")) || "", s = a ? n.rel(r, a) : void 0, l;
		if (s) {
			let e = s.target.startsWith("/") ? s.target.slice(1) : K(r, s.target), t = n.bytes(e);
			if (t) {
				let n = e.split(".").pop() || "png", r = n === "jpg" || n === "jpeg" ? "image/jpeg" : "image/" + n;
				l = {
					part: e,
					rId: a,
					url: URL.createObjectURL(new Blob([t], { type: r }))
				};
			}
		}
		return {
			id: t("sh"),
			kind: "pic",
			...c,
			image: l,
			name: o
		};
	}
	if (i === "sp" || i === "cxnSp") {
		let n = s && V(s, "prstGeom"), { fill: r, line: a } = ye(s);
		return {
			id: t("sh"),
			kind: i === "cxnSp" ? "cxnSp" : "sp",
			...c,
			preset: H(n, "prst") || (i === "cxnSp" ? "line" : "rect"),
			fill: r,
			line: a,
			name: o,
			text: _e(V(e, "txBody"))
		};
	}
	return {
		id: t("sh"),
		kind: "graphicFrame",
		...c,
		raw: W(e),
		name: o
	};
}
function K(e, t) {
	let n = e.includes("/") ? e.slice(0, e.lastIndexOf("/") + 1) : "", r = [];
	for (let e of (n + t).split("/")) e !== "." && e !== "" && (e === ".." ? r.pop() : r.push(e));
	return r.join("/");
}
async function xe(e) {
	let t = await R.open(e), n = t.xml("ppt/presentation.xml"), r = V(n, "sldSz"), i = {
		slideW: +(H(r, "cx") || 12192e3),
		slideH: +(H(r, "cy") || 6858e3),
		slides: [],
		format: t.variant("pptx")
	}, o = t.relsOf("ppt/presentation.xml"), s = V(n, "sldIdLst"), c = [], l = s ? B(s, "sldId") : [];
	for (let e of l) {
		let n = e.getAttributeNS(z, "id") || H(e, "r:id") || "", r = o.find((e) => e.id === n);
		if (!r) continue;
		let s = r.target.startsWith("/") ? r.target.slice(1) : K("ppt/presentation.xml", r.target);
		c.push(s);
		let l = a();
		if (l.part = s, t.has(s)) {
			let e = V(t.xml(s), "cSld"), n = e && V(e, "bg"), r = n && U(n, "srgbClr");
			r && (l.background = ("#" + H(r, "val")).toLowerCase());
			let i = e && V(e, "spTree");
			if (i) {
				for (let e of i.children) if ([
					"sp",
					"pic",
					"cxnSp",
					"graphicFrame"
				].includes(e.localName)) {
					let n = be(e, t, s);
					n && l.shapes.push(n);
				}
			}
			let a = t.relsOf(s).find((e) => e.type.endsWith("/notesSlide"));
			if (a) {
				let e = K(s, a.target);
				if (t.has(e)) {
					let n = U(t.xml(e), "txBody");
					n && (l.notes = B(n, "p").map((e) => B(e, "r").map((e) => V(e, "t")?.textContent || "").join("")).join("\n"));
				}
			}
		}
		i.slides.push(l);
	}
	return i.slides.length || (i.slides.push(a()), c.push("ppt/slides/slide1.xml")), {
		pres: i,
		pkg: t,
		slideParts: c
	};
}
// ../vpptx/src/engine/pptx-write.ts
var Se = M.p, q = M.a, J = `xmlns:a="${q}" xmlns:r="${M.r}" xmlns:p="${Se}"`, Y = (e) => e.replace("#", "").toUpperCase();
function Ce(e) {
	return e === null ? "<a:noFill/>" : e && e.color ? `<a:solidFill><a:srgbClr val="${Y(e.color)}"/></a:solidFill>` : "";
}
function we(e) {
	return e === null ? "<a:ln><a:noFill/></a:ln>" : e && e.color ? `<a:ln w="${Math.round((e.width || 1) * 12700)}"><a:solidFill><a:srgbClr val="${Y(e.color)}"/></a:solidFill>${e.dash ? `<a:prstDash val="${e.dash}"/>` : ""}</a:ln>` : "";
}
function Te(e) {
	if (!e) return "";
	let t = ["lang=\"en-US\""];
	e.size && t.push(`sz="${Math.round(e.size * 100)}"`), e.bold && t.push("b=\"1\""), e.italic && t.push("i=\"1\""), e.underline && t.push("u=\"sng\"");
	let n = "";
	return e.color && (n += `<a:solidFill><a:srgbClr val="${Y(e.color)}"/></a:solidFill>`), e.font && (n += `<a:latin typeface="${E(e.font)}"/>`), `<a:rPr ${t.join(" ")} dirty="0">${n}</a:rPr>`;
}
function Ee(e) {
	return `<a:p>${e.align || e.bullet ? `<a:pPr${e.align ? ` algn="${e.align === "center" ? "ctr" : e.align === "right" ? "r" : "l"}"` : ""}>${e.bullet ? "<a:buChar char=\"•\"/>" : "<a:buNone/>"}</a:pPr>` : ""}${e.runs.map((e) => `<a:r>${Te(e.rPr)}<a:t>${fe(e.text)}</a:t></a:r>`).join("") || "<a:endParaRPr lang=\"en-US\"/>"}</a:p>`;
}
function De(e) {
	return `<p:txBody><a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/>${(e && e.length ? e : [{ runs: [] }]).map(Ee).join("")}</p:txBody>`;
}
function X(e) {
	return `<a:xfrm${e.rot ? ` rot="${Math.round(e.rot * 6e4)}"` : ""}><a:off x="${Math.round(e.x)}" y="${Math.round(e.y)}"/><a:ext cx="${Math.round(e.w)}" cy="${Math.round(e.h)}"/></a:xfrm>`;
}
var Z = 1;
function Oe(e, t, n) {
	let r = ++Z;
	if (e.raw != null) return e.raw;
	if (e.kind === "pic" && e.image) {
		let t = e.image.rId;
		return `<p:pic><p:nvPicPr><p:cNvPr id="${r}" name="${E(e.name || "Picture")}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${t}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${X(e)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
	}
	let i = e.kind === "cxnSp" ? "p:cxnSp" : "p:sp", a = e.kind === "cxnSp" ? "p:nvCxnSpPr" : "p:nvSpPr", o = e.kind === "cxnSp" ? "p:cNvCxnSpPr" : "p:cNvSpPr", s = `<a:prstGeom prst="${E(e.preset || "rect")}"><a:avLst/></a:prstGeom>`, c = e.kind === "cxnSp" ? "" : De(e.text);
	return `<${i}><${a}><p:cNvPr id="${r}" name="${E(e.name || "Shape")}"/><${o}/><p:nvPr/></${a}><p:spPr>${X(e)}${s}${Ce(e.fill)}${we(e.line)}</p:spPr>${c}</${i}>`;
}
function ke(e, t, n) {
	Z = 1;
	let r = e.background ? `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${Y(e.background)}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` : "", i = e.shapes.map((e) => Oe(e, t, n)).join("");
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sld ${J}><p:cSld>${r}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` + i + "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>";
}
var Ae = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a:theme xmlns:a="${q}" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`, je = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldMaster ${J}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`, Me = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldLayout ${J} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`, Q = {
	pres: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
	master: "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
	layout: "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
	slide: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
	theme: "application/vnd.openxmlformats-officedocument.theme+xml"
};
async function Ne(e, t = {}) {
	let n = N[t.format || e.format || "pptx"] || N.pptx, r = t.pkg || ge("ppt/presentation.xml", "", n), i = t.slideParts && t.slideParts.length === e.slides.length ? t.slideParts : e.slides.map((e, t) => `ppt/slides/slide${t + 1}.xml`);
	for (let t = 0; t < e.slides.length; t++) {
		let n = i[t];
		for (let i of e.slides[t].shapes) if (i.kind === "pic" && i.image && i.image.url && (!i.image.part || !r.has(i.image.part))) {
			let e = await fetch(i.image.url), t = new Uint8Array(await e.arrayBuffer()), a = (i.image.part.split(".").pop() || "png").toLowerCase(), o = r.nextMediaName(a === "jpg" ? "jpeg" : a), s = "ppt/media/" + o;
			r.addPart(s, t), i.image.rId = r.addRel(n, M.rel + "/image", "../media/" + o), i.image.part = s;
		}
	}
	if (e.slides.forEach((e, n) => {
		let a = i[n];
		r.addPart(a, ke(e, r, a), Q.slide), t.pkg || r.addRel(a, M.rel + "/slideLayout", "../slideLayouts/slideLayout1.xml");
	}), !t.pkg) {
		r.addPart("ppt/theme/theme1.xml", Ae, Q.theme), r.addPart("ppt/slideMasters/slideMaster1.xml", je, Q.master), r.addPart("ppt/slideLayouts/slideLayout1.xml", Me, Q.layout), r.addRel("ppt/slideMasters/slideMaster1.xml", M.rel + "/slideLayout", "../slideLayouts/slideLayout1.xml"), r.addRel("ppt/slideMasters/slideMaster1.xml", M.rel + "/theme", "../theme/theme1.xml"), r.addRel("ppt/slideLayouts/slideLayout1.xml", M.rel + "/slideMaster", "../slideMasters/slideMaster1.xml");
		let t = e.slides.map((e, t) => `<p:sldId id="${256 + t}" r:id="rId${t + 2}"/>`).join("");
		r.setPart("ppt/presentation.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:presentation ${J}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${t}</p:sldIdLst><p:sldSz cx="${e.slideW}" cy="${e.slideH}"/><p:notesSz cx="${e.slideH}" cy="${e.slideW}"/></p:presentation>`);
		let n = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"];
		n.push(`<Relationship Id="rId1" Type="${M.rel}/slideMaster" Target="slideMasters/slideMaster1.xml"/>`), e.slides.forEach((e, t) => n.push(`<Relationship Id="rId${t + 2}" Type="${M.rel}/slide" Target="slides/slide${t + 1}.xml"/>`)), n.push("</Relationships>"), r.setPart("ppt/_rels/presentation.xml.rels", n.join("")), r.addRel("", M.rel + "/officeDocument", "ppt/presentation.xml");
	} else if (Pe(t.pkg) !== e.slides.length) {
		e.slides.map((e, t) => `<p:sldId id="${256 + t}" r:id="rId${t + 2}"/>`).join("");
		let t = r.xml("ppt/presentation.xml").children.find((e) => e.localName === "sldIdLst");
		t && (t.childNodes = [], t.children = []), r.setPart("ppt/presentation.xml", Fe(r, e));
		let n = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"];
		n.push(`<Relationship Id="rId1" Type="${M.rel}/slideMaster" Target="slideMasters/slideMaster1.xml"/>`), e.slides.forEach((e, t) => n.push(`<Relationship Id="rId${t + 2}" Type="${M.rel}/slide" Target="slides/slide${t + 1}.xml"/>`)), n.push("</Relationships>"), r.setPart("ppt/_rels/presentation.xml.rels", n.join(""));
	}
	return r.save(n.mime);
}
function Pe(e) {
	if (!e.has("ppt/presentation.xml")) return 0;
	let t = e.xml("ppt/presentation.xml").children.find((e) => e.localName === "sldIdLst");
	return t ? t.children.length : 0;
}
function Fe(e, t) {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:presentation ${J}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${t.slides.map((e, t) => `<p:sldId id="${256 + t}" r:id="rId${t + 2}"/>`).join("")}</p:sldIdLst><p:sldSz cx="${t.slideW}" cy="${t.slideH}"/><p:notesSz cx="${t.slideH}" cy="${t.slideW}"/></p:presentation>`;
}
// ../vpptx/src/engine/shapes.ts
function $(e, t, n, r, i, a) {
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
var Ie = [
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
], Le = (e) => e.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
function Re(e) {
	let t = r(e.x), n = r(e.y), i = r(e.w), a = r(e.h), o = `position:absolute;left:${t}px;top:${n}px;width:${i}px;height:${a}px;${e.rot ? `transform:rotate(${e.rot}deg);` : ""}`;
	if (e.raw != null) return `<div style="${o}border:1px dashed #999;background:#eee"></div>`;
	if (e.kind === "pic" && e.image) return `<div style="${o}"><img src="${e.image.url}" style="width:100%;height:100%;object-fit:fill" draggable="false"></div>`;
	let s = $(e.preset || "rect", i, a, e.fill ? e.fill.color : null, e.line ? e.line.color : null, e.line?.width || 1), c = (e.text || []).map((e) => {
		let t = e.align || "left", n = e.runs.map((e) => {
			let t = [];
			return e.rPr?.bold && t.push("font-weight:bold"), e.rPr?.italic && t.push("font-style:italic"), e.rPr?.underline && t.push("text-decoration:underline"), e.rPr?.size && t.push(`font-size:${e.rPr.size}px`), e.rPr?.color && t.push(`color:${e.rPr.color}`), e.rPr?.font && t.push(`font-family:${e.rPr.font},sans-serif`), `<span style="${t.join(";")}">${Le(e.text)}</span>`;
		}).join("");
		return `<div style="text-align:${t};min-height:1em">${e.bullet ? "• " : ""}${n}</div>`;
	}).join(""), l = (e.text || []).some((e) => e.runs.some((e) => e.text));
	return `<div style="${o}"><svg width="${i}" height="${a}" viewBox="0 0 ${i} ${a}" style="position:absolute;inset:0;overflow:visible">${s}</svg>` + (l || e.kind === "sp" ? `<div style="position:absolute;inset:4px;overflow:hidden;color:#000;font-size:18px">${c}</div>` : "") + "</div>";
}
function ze(e, t, n) {
	let i = n / r(t.slideW), a = e.shapes.map(Re).join("");
	return `<div style="position:relative;width:${r(t.slideW)}px;height:${r(t.slideH)}px;background:${e.background || "#fff"};transform:scale(${i});transform-origin:top left">${a}</div>`;
}
export { xe as a, i as c, s as d, r as f, Ne as i, o as l, Ie as n, n as o, t as p, $ as r, c as s, ze as t, a as u };
