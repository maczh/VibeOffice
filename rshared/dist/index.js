import { Component as e, createElement as t, useCallback as n, useEffect as r, useMemo as i, useRef as a, useState as o } from "react";
import { Fragment as s, jsx as c, jsxs as l } from "react/jsx-runtime";
// src/online/types.ts
var u = class extends Error {
	status;
	constructor(e, t) {
		super(e), this.status = t;
	}
};
function d(e) {
	try {
		let t = new URL(e).pathname.split("/").filter(Boolean).pop() || "document";
		return decodeURIComponent(t);
	} catch {
		return "document";
	}
}
function f(e) {
	let t = {};
	return e.token ? t.Authorization = "Bearer " + e.token : e.username != null && (t.Authorization = "Basic " + btoa(unescape(encodeURIComponent(`${e.username}:${e.password ?? ""}`)))), t;
}
// src/core/xml.ts
var p = {
	lt: "<",
	gt: ">",
	amp: "&",
	quot: "\"",
	apos: "'"
};
function m(e) {
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
		return p[t] == null ? e : p[t];
	});
}
var h = (e) => (e = String(e).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"), /[\t\n\r]/.test(e) ? e.replace(/\t/g, "&#9;").replace(/\n/g, "&#10;").replace(/\r/g, "&#13;") : e), g = /* @__PURE__ */ RegExp("[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F\\uFFFE\\uFFFF]", "g"), _ = /* @__PURE__ */ RegExp("[\\x00-\\x08\\x0B-\\x1F\\uFFFE\\uFFFF]", "g"), v = (e, t) => String(e).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/_x([0-9A-Fa-f]{4})_/g, "_x005F_x$1_").replace(t ? g : _, (e) => "_x" + e.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0") + "_"), ee = (e) => e.indexOf("_x") < 0 ? e : e.replace(/_x([0-9A-Fa-f]{4})_/g, (e, t) => String.fromCharCode(parseInt(t, 16))), y = class {
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
}, b = class e {
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
		this.children = [], this.childNodes = [new y(String(e), this)];
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
		if (this.attrs && (n.attrs = { ...this.attrs }), t) for (let e of this.childNodes) n.appendChild(e.nodeType === 3 ? new y(e.data, n) : e.cloneNode(!0));
		return n;
	}
}, x = /([^\s=\/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
function te(e, t) {
	x.lastIndex = 0;
	let n;
	for (; n = x.exec(e);) {
		let e = t.attrs ??= Object.create(null), r = n[2] === void 0 ? n[3] : n[2];
		e[n[1]] = r.indexOf("&") < 0 ? r : m(r);
	}
}
function ne(e) {
	let t = Object.create(null);
	x.lastIndex = 0;
	let n;
	for (; n = x.exec(e);) {
		let e = n[2] === void 0 ? n[3] : n[2];
		t[n[1]] = e.indexOf("&") < 0 ? e : m(e);
	}
	return t;
}
var re = /* @__PURE__ */ new Set([
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
function S(e, t) {
	e && e.charCodeAt(0) === 65279 && (e = e.slice(1));
	let n = !!t?.keepWhitespace, r = !!t?.strict, i = new b("#document", null), a = i, o = 0, s = e.length, c = (e) => {
		a !== i && (n || !/^\s*$/.test(e) || re.has(a.localName)) && a.childNodes.push(new y(m(e), a));
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
				a !== i && a.childNodes.push(new y(r, a)), o = n < 0 ? s : n + 3;
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
			let c = e.slice(t + 1, i), l = e.charCodeAt(n - 1) === 47, u = new b(c, a), d = l ? n - 1 : n;
			i < d && te(e.slice(i, d), u), a.childNodes.push(u), a.children.push(u), l || (a = u), o = n + 1;
		}
	}
	let l = i.children[0];
	if (!l) throw Error("XML parse error: no root element");
	if (r && (a !== i || i.children.length !== 1)) throw Error("XML parse error: incomplete document");
	return l;
}
function ie(e) {
	return e.indexOf("\r") >= 0 && (e = e.replace(/\r\n?/g, "\n")), e.indexOf("<![CDATA[") < 0 ? m(e) : e.split(/(<!\[CDATA\[[\s\S]*?\]\]>)/).map((e) => e.startsWith("<![CDATA[") ? e.slice(9, -3) : m(e)).join("");
}
function ae(e, t, n, r) {
	let i = /<(\/?)(?:[A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)([^>]*?)(\/?)>/g;
	i.lastIndex = t;
	let a = t, o;
	for (; (o = i.exec(e)) && o.index < n;) {
		let t = o.index > a ? e.slice(a, o.index) : "";
		r(o[2], o[1] ? 2 : o[4] ? 3 : 1, o[3], t), a = i.lastIndex;
	}
}
function C(e) {
	return e.length >= 2 && e[0] === 255 && e[1] === 254 ? new TextDecoder("utf-16le").decode(e.subarray(2)) : e.length >= 2 && e[0] === 254 && e[1] === 255 ? new TextDecoder("utf-16be").decode(e.subarray(2)) : e.length >= 4 && e[0] === 0 && e[1] === 60 && e[2] === 0 && e[3] === 63 ? new TextDecoder("utf-16be").decode(e) : e.length >= 4 && e[0] === 60 && e[1] === 0 && e[2] === 63 && e[3] === 0 ? new TextDecoder("utf-16le").decode(e) : new TextDecoder("utf-8").decode(e);
}
function w(e) {
	if (!e) return "";
	let t = {};
	for (let n = e; n && n.nodeType === 1; n = n.parentNode) if (n.attrs) for (let e in n.attrs) (e === "xmlns" || e.startsWith("xmlns:")) && (t[e] = n.attrs[e]);
	let n = (e, r) => {
		if (e.nodeType === 3) return e.data.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
		let i = e, a = i.attrs ? { ...i.attrs } : {};
		if (r) for (let [e, n] of Object.entries(t)) e in a || (a[e] = n);
		let o = "<" + i.nodeName + Object.entries(a).map(([e, t]) => " " + e + "=\"" + h(t) + "\"").join("");
		return i.childNodes.length ? o + ">" + i.childNodes.map((e) => n(e, !1)).join("") + "</" + i.nodeName + ">" : o + "/>";
	};
	return n(e, !0);
}
function oe(e, t, n) {
	let r = new b(e, null);
	if (t) for (let [e, n] of Object.entries(t)) r.setAttribute(e, n);
	return n != null && (r.textContent = n), r;
}
// src/online/s3.ts
var T = new TextEncoder(), E = (e) => Array.from(new Uint8Array(e), (e) => e.toString(16).padStart(2, "0")).join(""), se = async (e) => E(await crypto.subtle.digest("SHA-256", typeof e == "string" ? T.encode(e) : e)), D = async (e, t) => {
	let n = await crypto.subtle.importKey("raw", e, {
		name: "HMAC",
		hash: "SHA-256"
	}, !1, ["sign"]);
	return crypto.subtle.sign("HMAC", n, T.encode(t));
}, O = (e, t = !1) => encodeURIComponent(e).replace(/[!'()*]/g, (e) => "%" + e.charCodeAt(0).toString(16).toUpperCase()).replace(/%2F/g, t ? "/" : "%2F");
function k(e) {
	let t = new URL(e.endpoint), n = e.key.replace(/^\/+/, "");
	return e.forcePathStyle ? t.pathname = "/" + e.bucket + "/" + n : (t.hostname = e.bucket + "." + t.hostname, t.pathname = "/" + n), t;
}
async function A(e, t, n, r, i = {}) {
	let a = e.region || "us-east-1", o = (/* @__PURE__ */ new Date()).toISOString().replace(/[:-]|\.\d{3}/g, ""), s = o.slice(0, 8);
	for (let [e, n] of Object.entries(i)) t.searchParams.set(e, n);
	let c = await se(r ?? ""), l = {
		host: t.host,
		"x-amz-content-sha256": c,
		"x-amz-date": o
	};
	e.sessionToken && (l["x-amz-security-token"] = e.sessionToken);
	let u = Object.keys(l).sort().join(";"), d = Object.keys(l).sort().map((e) => `${e}:${l[e].trim()}\n`).join(""), f = [...t.searchParams.entries()].sort(([e], [t]) => e < t ? -1 : +(e > t)).map(([e, t]) => `${O(e)}=${O(t)}`).join("&"), p = [
		n,
		O(t.pathname, !0),
		f,
		d,
		u,
		c
	].join("\n"), m = `${s}/${a}/s3/aws4_request`, h = [
		"AWS4-HMAC-SHA256",
		o,
		m,
		await se(p)
	].join("\n"), g = E(await D(await D(await D(await D(await D(T.encode("AWS4" + e.secretAccessKey), s), a), "s3"), "aws4_request"), h)), _ = `AWS4-HMAC-SHA256 Credential=${e.accessKeyId}/${m}, SignedHeaders=${u}, Signature=${g}`, v = {
		...l,
		Authorization: _
	};
	return delete v.host, fetch(t.toString(), {
		method: n,
		headers: v,
		body: r
	});
}
async function ce(e) {
	let t = await A(e, k(e), "GET", null);
	if (!t.ok) throw new u(`S3 GET failed: ${t.status} ${t.statusText}`, t.status);
	return {
		blob: await t.blob(),
		name: d(k(e).toString())
	};
}
async function le(e, t) {
	let n = new Uint8Array(await t.arrayBuffer()), r = await A(e, k(e), "PUT", n);
	if (!r.ok) throw new u(`S3 PUT failed: ${r.status} ${r.statusText}`, r.status);
}
// src/online/webdav.ts
async function ue(e) {
	let t = await fetch(e.url, { headers: f(e) });
	if (!t.ok) throw new u(`WebDAV GET failed: ${t.status} ${t.statusText}`, t.status);
	return {
		blob: await t.blob(),
		name: d(e.url)
	};
}
async function de(e, t) {
	let n = await fetch(e.url, {
		method: "PUT",
		headers: {
			...f(e),
			"Content-Type": t.type || "application/octet-stream"
		},
		body: t
	});
	if (!n.ok) throw new u(`WebDAV PUT failed: ${n.status} ${n.statusText}`, n.status);
}
// src/online/http.ts
async function fe(e) {
	let t = await fetch(e.url, { headers: e.headers });
	if (!t.ok) throw new u(`Download failed: ${t.status} ${t.statusText}`, t.status);
	let n = t.headers.get("Content-Disposition") || "", r = /filename\*?=(?:UTF-8''|")?([^";]+)/i.exec(n);
	return {
		blob: await t.blob(),
		name: r ? decodeURIComponent(r[1].replace(/"$/, "")) : d(e.url)
	};
}
async function pe(e, t, n) {
	let r = new FormData();
	for (let [t, n] of Object.entries(e.fields || {})) r.append(t, n);
	r.append(e.field || "file", t, n);
	let i = await fetch(e.url, {
		method: "POST",
		headers: e.headers,
		body: r
	});
	if (!i.ok) throw new u(`Upload failed: ${i.status} ${i.statusText}`, i.status);
}
// src/core/zip.ts
var j = new TextEncoder(), M = new TextDecoder("utf-8"), me = (() => {
	let e = /* @__PURE__ */ new Uint32Array(256);
	for (let t = 0; t < 256; t++) {
		let n = t;
		for (let e = 0; e < 8; e++) n = n & 1 ? 3988292384 ^ n >>> 1 : n >>> 1;
		e[t] = n >>> 0;
	}
	return e;
})();
function N(e) {
	let t = 4294967295;
	for (let n = 0; n < e.length; n++) t = me[(t ^ e[n]) & 255] ^ t >>> 8;
	return (t ^ 4294967295) >>> 0;
}
var he = [
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
], ge = [
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
], _e = [
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
], ve = [
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
], ye = [
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
function P(e, t) {
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
var F = null, I = null;
function be() {
	if (F) return;
	let e = /* @__PURE__ */ new Uint8Array(288);
	for (let t = 0; t < 144; t++) e[t] = 8;
	for (let t = 144; t < 256; t++) e[t] = 9;
	for (let t = 256; t < 280; t++) e[t] = 7;
	for (let t = 280; t < 288; t++) e[t] = 8;
	F = P(e, 288), I = P((/* @__PURE__ */ new Uint8Array(30)).fill(5), 30);
}
function L(e, t) {
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
				let e = he[n] + c(ge[n]), r = l(t), i = _e[r] + c(ve[r]);
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
		} else if (t === 1) be(), u(F, I);
		else if (t === 2) {
			let e = c(5) + 257, t = c(5) + 1, n = c(4) + 4, r = /* @__PURE__ */ new Uint8Array(19);
			for (let e = 0; e < n; e++) r[ye[e]] = c(3);
			let i = P(r, 19), a = new Uint8Array(e + t), o = 0;
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
			u(P(a.subarray(0, e), e), P(a.subarray(e), t));
		} else throw Error("inflate: bad block type");
	} while (!d);
	return a.slice(0, o);
}
async function R(e, t) {
	let n = new Response(new Blob([e]).stream().pipeThrough(t));
	return new Uint8Array(await n.arrayBuffer());
}
async function z(e, t) {
	if (typeof DecompressionStream == "function") try {
		return await R(e, new DecompressionStream("deflate-raw"));
	} catch {}
	return L(e, t);
}
async function xe(e) {
	if (typeof CompressionStream == "function") try {
		return await R(e, new CompressionStream("deflate-raw"));
	} catch {}
	return null;
}
function B(e, t, n, r, i, a, o, s, c) {
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
			else if (r === 8) d = await z(u, a);
			else throw Error("Unsupported compression method " + r + " in " + n);
			return l = d, d;
		},
		async text() {
			return M.decode(await this.bytes());
		}
	};
}
var V = (e, t) => t & 2048 ? M.decode(e) : Array.from(e, (e) => String.fromCharCode(e)).join("");
function H(e, t) {
	if (t.name.endsWith("/")) return;
	e.set(t.name, t);
	let n = t.name.replace(/\\/g, "/");
	n !== t.name && !e.has(n) && e.set(n, t);
}
function Se(e, t) {
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
		let n = t.getUint16(o + 8, !0), r = t.getUint16(o + 10, !0), i = t.getUint32(o + 16, !0), s = t.getUint32(o + 20, !0), c = t.getUint32(o + 24, !0), l = t.getUint16(o + 28, !0), u = t.getUint16(o + 30, !0), d = t.getUint16(o + 32, !0), f = t.getUint32(o + 42, !0), p = V(e.subarray(o + 46, o + 46 + l), n), m = o + 46 + l, h = Math.min(e.length, m + u);
		for (; m + 4 <= h;) {
			let e = t.getUint16(m, !0), n = t.getUint16(m + 2, !0);
			if (e === 1) {
				let e = m + 4;
				c === 4294967295 && (c = Number(t.getBigUint64(e, !0)), e += 8), s === 4294967295 && (s = Number(t.getBigUint64(e, !0)), e += 8), f === 4294967295 && (f = Number(t.getBigUint64(e, !0)));
			}
			m += 4 + n;
		}
		o += 46 + l + u + d, H(a, B(e, t, p, r, s, c, f, void 0, i));
	}
	return a.size ? a : null;
}
function Ce(e, t) {
	let n = /* @__PURE__ */ new Map(), r = 0;
	for (; r + 30 <= e.length;) {
		if (t.getUint32(r, !0) !== 67324752) {
			let t = r + 1;
			for (; t + 4 <= e.length && (e[t] !== 80 || e[t + 1] !== 75 || e[t + 2] !== 3 || e[t + 3] !== 4);) t++;
			if (t + 30 > e.length) break;
			r = t;
			continue;
		}
		let i = t.getUint16(r + 6, !0), a = t.getUint16(r + 8, !0), o = t.getUint32(r + 18, !0), s = t.getUint32(r + 22, !0), c = t.getUint16(r + 26, !0), l = t.getUint16(r + 28, !0), u = V(e.subarray(r + 30, r + 30 + c), i), d = r + 30 + c + l, f = d + o;
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
		H(n, B(e, t, u, a, o, s, r, d)), r = Math.max(f, r + 30);
	}
	return n;
}
async function U(e) {
	let t = e instanceof Uint8Array ? e : new Uint8Array(e), n = new DataView(t.buffer, t.byteOffset, t.byteLength);
	if (t.length < 22) throw Error("This file is not a valid Office Open XML (ZIP) package.");
	if (n.getUint32(0, !0) !== 67324752 && n.getUint32(0, !0) === 3759263696) {
		let e = new TextDecoder("utf-16le").decode(t.subarray(0, Math.min(t.length, 1 << 20))), n = /EncryptedPackage/.test(e), r = /* @__PURE__ */ Error(n ? "This file is password-protected." : "This is a binary Office 97–2003 file, not an Office Open XML package.");
		throw r.code = "ole", r.encrypted = n, r.bytes = t, r;
	}
	let r = null;
	try {
		r = Se(t, n);
	} catch {
		r = null;
	}
	if (!r) {
		try {
			r = Ce(t, n);
		} catch {
			r = null;
		}
		if (!r || !r.size) throw Error(n.getUint32(0, !0) === 67324752 ? "This file is damaged and could not be repaired." : "This file is not a valid Office Open XML (ZIP) package.");
		r.repaired = !0;
	}
	return r;
}
function we(e) {
	return {
		time: e.getHours() << 11 | e.getMinutes() << 5 | Math.floor(e.getSeconds() / 2),
		date: e.getFullYear() - 1980 << 9 | e.getMonth() + 1 << 5 | e.getDate()
	};
}
async function W(e, t) {
	let n = [], r = [], i = 0, { time: a, date: o } = we(/* @__PURE__ */ new Date());
	for (let t of e) {
		let e = j.encode(t.name), s = typeof t.data == "string" ? j.encode(t.data) : t.data, c = N(s), l = null;
		!t.store && s.length > 64 && (l = await xe(s));
		let u = l && l.length < s.length ? 8 : 0, d = u === 8 ? l : s, f = new Uint8Array(30 + e.length), p = new DataView(f.buffer);
		p.setUint32(0, 67324752, !0), p.setUint16(4, 20, !0), p.setUint16(6, 2048, !0), p.setUint16(8, u, !0), p.setUint16(10, a, !0), p.setUint16(12, o, !0), p.setUint32(14, c, !0), p.setUint32(18, d.length, !0), p.setUint32(22, s.length, !0), p.setUint16(26, e.length, !0), p.setUint16(28, 0, !0), f.set(e, 30), n.push(f, d);
		let m = new Uint8Array(46 + e.length), h = new DataView(m.buffer);
		h.setUint32(0, 33639248, !0), h.setUint16(4, 20, !0), h.setUint16(6, 20, !0), h.setUint16(8, 2048, !0), h.setUint16(10, u, !0), h.setUint16(12, a, !0), h.setUint16(14, o, !0), h.setUint32(16, c, !0), h.setUint32(20, d.length, !0), h.setUint32(24, s.length, !0), h.setUint16(28, e.length, !0), h.setUint32(42, i, !0), m.set(e, 46), r.push(m), i += f.length + d.length;
	}
	let s = r.reduce((e, t) => e + t.length, 0), c = /* @__PURE__ */ new Uint8Array(22), l = new DataView(c.buffer);
	return l.setUint32(0, 101010256, !0), l.setUint16(8, e.length, !0), l.setUint16(10, e.length, !0), l.setUint32(12, s, !0), l.setUint32(16, i, !0), new Blob([
		...n,
		...r,
		c
	], { type: t || "application/zip" });
}
var Te = (e) => j.encode(e), Ee = (e) => M.decode(e), G = Object.freeze({
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
}), K = Object.freeze(Object.fromEntries([
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
}))), De = () => {
	if (crypto.randomUUID) return crypto.randomUUID();
	let e = crypto.getRandomValues(/* @__PURE__ */ new Uint8Array(16));
	e[6] = e[6] & 15 | 64, e[8] = e[8] & 63 | 128;
	let t = Array.from(e, (e) => e.toString(16).padStart(2, "0")).join("");
	return `${t.slice(0, 8)}-${t.slice(8, 12)}-${t.slice(12, 16)}-${t.slice(16, 20)}-${t.slice(20)}`;
}, q = "[Content_Types].xml", J = (e) => {
	let t = e.lastIndexOf("/");
	return (t < 0 ? "" : e.slice(0, t + 1)) + "_rels/" + (t < 0 ? e : e.slice(t + 1)) + ".rels";
};
function Y(e, t) {
	if (t.startsWith("/")) return t.slice(1);
	let n = ((e.includes("/") ? e.slice(0, e.lastIndexOf("/") + 1) : "") + t).split("/"), r = [];
	for (let e of n) e !== "." && e !== "" && (e === ".." ? r.pop() : r.push(e));
	return r.join("/");
}
var X = (e) => e.replace(G.strictRel + "/", G.rel + "/"), Oe = class e {
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
		let n = await U(t), r = new e();
		r.repaired = !!n.repaired;
		for (let [e, t] of n) r.parts.set(e, await t.bytes());
		let i = r.parts.get(q);
		if (!i) throw Error("This package has no [Content_Types].xml part.");
		let a = S(C(i));
		for (let e of a.children) e.localName === "Default" ? r.defaults.set(e.getAttribute("Extension") || "", e.getAttribute("ContentType") || "") : e.localName === "Override" && r.overrides.set((e.getAttribute("PartName") || "").replace(/^\//, ""), e.getAttribute("ContentType") || "");
		let o = r.relsOf("").find((e) => X(e.type) === G.rel + "/officeDocument");
		return o && (r.main = Y("", o.target)), r;
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
			t = C(n), this.textCache.set(e, t);
		}
		return t;
	}
	xml(e) {
		let t = this.xmlCache.get(e);
		return t || (t = S(this.text(e)), this.xmlCache.set(e, t)), t;
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
		let r = this.parts.get(J(e));
		if (r) {
			let e = S(C(r));
			for (let t of e.children) t.localName === "Relationship" && n.push({
				id: t.getAttribute("Id") || "",
				type: X(t.getAttribute("Type") || ""),
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
		return this.relsOf(e).filter((e) => e.type === t && e.mode !== "External").map((t) => Y(e, t.target));
	}
	setPart(e, t) {
		this.parts.set(e, typeof t == "string" ? new TextEncoder().encode(t) : t), this.edited.add(e), this.textCache.delete(e), this.xmlCache.delete(e);
	}
	addPart(e, t, n) {
		if (this.setPart(e, t), n) this.overrides.set(e, n), this.ctDirty = !0;
		else {
			let t = e.includes(".") ? e.slice(e.lastIndexOf(".") + 1) : "";
			t && !this.defaults.has(t) && (this.defaults.set(t, ke(t)), this.ctDirty = !0);
		}
	}
	removePart(e) {
		this.parts.delete(e), this.edited.delete(e), this.textCache.delete(e), this.xmlCache.delete(e), this.overrides.delete(e) && (this.ctDirty = !0);
		let t = J(e);
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
		return Object.values(K).find((e) => e.contentType === t)?.ext || e;
	}
	writeContentTypes() {
		if (!this.ctDirty) return;
		let e = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", "<Types xmlns=\"" + G.ct + "\">"];
		for (let [t, n] of [...this.defaults].sort()) e.push(`<Default Extension="${h(t)}" ContentType="${h(n)}"/>`);
		for (let [t, n] of [...this.overrides].sort()) e.push(`<Override PartName="/${h(t)}" ContentType="${h(n)}"/>`);
		e.push("</Types>"), this.parts.set(q, new TextEncoder().encode(e.join("")));
	}
	writeRels() {
		for (let e of this.relsDirty) {
			let t = this.relsCache.get(e) || [], n = e, r = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", "<Relationships xmlns=\"" + G.pkg + "\">"];
			for (let e of t) r.push(`<Relationship Id="${h(e.id)}" Type="${h(e.type)}" Target="${h(e.target)}"${e.mode ? ` TargetMode="${h(e.mode)}"` : ""}/>`);
			r.push("</Relationships>");
			let i = J(n);
			this.parts.set(i, new TextEncoder().encode(r.join(""))), this.overrides.has(i) || (this.defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml"), this.ctDirty = !0);
		}
	}
	async save(e) {
		return this.writeRels(), this.writeContentTypes(), W([...this.parts].map(([e, t]) => ({
			name: e,
			data: t
		})), e || "application/zip");
	}
};
function ke(e) {
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
function Ae(e, t, n, r = {}) {
	let i = new Oe();
	i.main = e, i.defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml"), i.defaults.set("xml", "application/xml");
	for (let [e, t] of Object.entries(r)) i.defaults.set(e, t);
	return i.overrides.set(e, n.contentType), i.setPart(e, t), i.addRel("", G.rel + "/officeDocument", e), i;
}
function je(e) {
	return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" + w(e);
}
// src/chrome.tsx
function Me({ menus: e }) {
	let [t, n] = o(null);
	return r(() => {
		let e = () => n(null);
		return document.addEventListener("mousedown", e), () => document.removeEventListener("mousedown", e);
	}, []), /* @__PURE__ */ c("nav", {
		className: "vo-menubar",
		"aria-label": "Application menu",
		children: e.map((e) => /* @__PURE__ */ l("div", {
			className: "vo-menu",
			children: [/* @__PURE__ */ c("button", {
				onMouseDown: (r) => {
					r.stopPropagation(), n(t === e.label ? null : e.label);
				},
				children: e.label
			}), t === e.label && /* @__PURE__ */ c("div", {
				className: "vo-menu-pop",
				role: "menu",
				children: e.items.map((t, r) => t.sep ? /* @__PURE__ */ c("div", { className: "vo-menu-sep" }, `${e.label}-sep-${r}`) : /* @__PURE__ */ l("button", {
					disabled: t.disabled,
					role: "menuitem",
					onMouseDown: (e) => {
						e.stopPropagation(), n(null), t.run?.();
					},
					children: [
						t.checked ? "✓ " : "",
						t.label,
						t.key && /* @__PURE__ */ c("span", {
							className: "vo-menu-key",
							children: t.key
						})
					]
				}, t.label))
			})]
		}, e.label))
	});
}
var Ne = {
	new: "＋",
	open: "📂",
	save: "💾",
	undo: "↶",
	redo: "↷",
	print: "🖨",
	find: "🔍",
	cut: "✂",
	copy: "▣",
	paste: "📋",
	bold: "B",
	italic: "I",
	underline: "U",
	table: "▦",
	break: "↵",
	date: "📅"
};
function Pe({ items: e }) {
	return /* @__PURE__ */ c("div", {
		className: "vo-toolbar vo-commandbar",
		children: e.map((e) => e.kind === "sep" ? /* @__PURE__ */ c("span", { className: "vo-tool-sep" }, e.id) : /* @__PURE__ */ c("button", {
			className: "vo-tool-button",
			title: e.title ?? e.label,
			"aria-label": e.title ?? e.label,
			onClick: () => void e.run?.(),
			children: e.icon ?? /* @__PURE__ */ c("span", {
				className: `vo-tool-icon vo-tool-icon-${e.id}`,
				"aria-hidden": "true",
				children: Ne[e.id] ?? "◆"
			})
		}, e.id))
	});
}
function Fe({ tabs: e }) {
	let [t, n] = o(e[0]?.id), r = e.find((e) => e.id === t) ?? e[0];
	return /* @__PURE__ */ l("section", {
		className: "vo-ribbon",
		children: [/* @__PURE__ */ c("div", {
			className: "vo-ribbon-tabs",
			children: e.map((e) => /* @__PURE__ */ c("button", {
				className: e.id === t ? "active" : "",
				onClick: () => n(e.id),
				children: e.label
			}, e.id))
		}), r && /* @__PURE__ */ c("div", {
			className: "vo-ribbon-items",
			children: r.items.map((e) => /* @__PURE__ */ l("button", {
				className: "vo-ribbon-item",
				title: e.title ?? e.label,
				"aria-label": e.title ?? e.label,
				onClick: () => void e.run?.(),
				children: [e.icon ?? /* @__PURE__ */ c("span", {
					className: "vo-tool-icon",
					"aria-hidden": "true",
					children: Ne[e.id] ?? "◆"
				}), /* @__PURE__ */ c("span", { children: e.label ?? e.id })]
			}, e.id))
		})]
	});
}
function Ie({ children: e }) {
	return /* @__PURE__ */ c("div", {
		className: "vo-status",
		children: e
	});
}
function Z({ title: e, children: t, onClose: n }) {
	return /* @__PURE__ */ c("div", {
		className: "vo-dialog-backdrop",
		role: "presentation",
		children: /* @__PURE__ */ l("section", {
			className: "vo-dialog",
			role: "dialog",
			"aria-modal": "true",
			"aria-label": e,
			children: [/* @__PURE__ */ l("header", { children: [e, /* @__PURE__ */ c("button", {
				"aria-label": "Close",
				onClick: n,
				children: "×"
			})] }), /* @__PURE__ */ c("div", {
				className: "vo-dialog-body",
				children: t
			})]
		})
	});
}
function Le({ mode: e, fileName: t, onSubmit: n, onClose: r }) {
	let [i, a] = o("");
	return /* @__PURE__ */ l(Z, {
		title: e === "open" ? "Open from Online" : "Save to Online",
		onClose: r,
		children: [
			/* @__PURE__ */ l("label", { children: ["HTTP/WebDAV/Presigned URL", /* @__PURE__ */ c("input", {
				autoFocus: !0,
				value: i,
				onChange: (e) => a(e.target.value),
				placeholder: "https://…"
			})] }),
			t && /* @__PURE__ */ c("p", { children: t }),
			/* @__PURE__ */ l("div", { children: [/* @__PURE__ */ c("button", {
				onClick: r,
				children: "Cancel"
			}), /* @__PURE__ */ c("button", {
				disabled: !i.trim(),
				onClick: () => n(i.trim()),
				children: e === "open" ? "Open" : "Save"
			})] })
		]
	});
}
function Re({ x: e, y: t, items: n, onClose: i }) {
	let s = a(null), [u, d] = o(null);
	return r(() => {
		let e = (e) => {
			s.current?.contains(e.target) || i();
		}, t = (e) => {
			e.key === "Escape" && i();
		};
		return document.addEventListener("mousedown", e, !0), document.addEventListener("keydown", t), s.current?.focus(), () => {
			document.removeEventListener("mousedown", e, !0), document.removeEventListener("keydown", t);
		};
	}, [i]), /* @__PURE__ */ c("div", {
		ref: s,
		className: "vo-context",
		style: {
			left: e,
			top: t
		},
		role: "menu",
		tabIndex: -1,
		children: n.map((n, r) => n.sep ? /* @__PURE__ */ c("div", { className: "vo-menu-sep" }, `sep-${r}`) : /* @__PURE__ */ l("div", {
			className: "vo-context-row",
			children: [/* @__PURE__ */ l("button", {
				disabled: n.disabled,
				role: "menuitem",
				onMouseEnter: () => d(n.label),
				onClick: () => {
					n.sub?.length || (i(), n.run?.());
				},
				children: [
					n.checked ? "✓ " : "",
					n.label,
					n.key && /* @__PURE__ */ c("span", {
						className: "vo-menu-key",
						children: n.key
					}),
					n.sub?.length ? " ▶" : ""
				]
			}), u === n.label && n.sub?.length ? /* @__PURE__ */ c("div", {
				className: "vo-context-sub",
				children: /* @__PURE__ */ c(Re, {
					x: e + 180,
					y: t + r * 28,
					items: n.sub,
					onClose: i
				})
			}) : null]
		}, `${n.label}-${r}`))
	});
}
// src/common.tsx
var Q = "vibeoffice.recent";
function ze() {
	try {
		let e = JSON.parse(localStorage.getItem(Q) || "[]");
		return Array.isArray(e) ? e.slice(0, 12) : [];
	} catch {
		return [];
	}
}
function Be() {
	let [e, t] = o(ze);
	return {
		recent: e,
		remember: n((e) => t((t) => {
			let n = [{
				...e,
				openedAt: Date.now()
			}, ...t.filter((t) => t.name !== e.name || t.app !== e.app)].slice(0, 12);
			try {
				localStorage.setItem(Q, JSON.stringify(n));
			} catch {}
			return n;
		}), []),
		remove: n((e) => t((t) => {
			let n = t.filter((t) => t !== e);
			try {
				localStorage.setItem(Q, JSON.stringify(n));
			} catch {}
			return n;
		}), [])
	};
}
function Ve(e) {
	let [t, n] = o(() => e || localStorage.getItem("vibeoffice.look") || "classic");
	return r(() => {
		try {
			localStorage.setItem("vibeoffice.look", t);
		} catch {}
	}, [t]), {
		look: t,
		setLook: n,
		layout: t === "paper2016" ? "ribbon" : "bars"
	};
}
function He(e, t = "You have unsaved changes.") {
	r(() => {
		let n = (n) => {
			e && (n.preventDefault(), n.returnValue = t);
		};
		return window.addEventListener("beforeunload", n), () => window.removeEventListener("beforeunload", n);
	}, [e, t]);
}
function Ue() {
	let [e, t] = o(() => localStorage.getItem("vibeoffice.lang") || "en");
	return {
		lang: e,
		setLanguage: n((e) => {
			t(e);
			try {
				localStorage.setItem("vibeoffice.lang", e);
			} catch {}
			document.documentElement.lang = e;
		}, []),
		t: (e, ...t) => e.replace(/\{(\d+)\}/g, (e, n) => t[+n] == null ? e : String(t[+n]))
	};
}
function We(e = "/sw.js") {
	let [t, i] = o(!1), [a, s] = o(null);
	return r(() => {
		if (!("serviceWorker" in navigator)) return;
		let t = !0;
		return navigator.serviceWorker.register(e).then((e) => {
			t && s(e), e.addEventListener("updatefound", () => e.installing?.addEventListener("statechange", () => {
				e.installing?.state === "installed" && navigator.serviceWorker.controller && i(!0);
			}));
		}).catch(() => {}), () => {
			t = !1;
		};
	}, [e]), {
		updateAvailable: t,
		applyUpdate: n(() => {
			a?.waiting?.postMessage({ type: "SKIP_WAITING" }), window.location.reload();
		}, [a])
	};
}
function Ge() {
	let [e, t] = o({
		visible: !1,
		prompt: "",
		choices: []
	}), r = n((e = "", n = []) => t({
		visible: !0,
		prompt: e,
		choices: n
	}), []), i = n(() => t((e) => ({
		...e,
		visible: !1
	})), []);
	return {
		...e,
		show: r,
		hide: i,
		choose: (e) => {
			e.run?.(), i();
		}
	};
}
function Ke() {
	let { lang: e, setLanguage: t } = Ue();
	return /* @__PURE__ */ l("label", {
		className: "vo-language-select",
		children: [/* @__PURE__ */ c("span", { children: "Language" }), /* @__PURE__ */ c("select", {
			value: e,
			onChange: (e) => t(e.target.value),
			children: [
				"en",
				"ja",
				"ko",
				"fr",
				"de",
				"zh-CN",
				"zh-Hant",
				"qps"
			].map((e) => /* @__PURE__ */ c("option", { children: e }, e))
		})]
	});
}
function qe({ accept: e, onFiles: t, children: n }) {
	let r = (t) => !e || e.split(",").some((e) => e.trim().startsWith(".") && t.name.toLowerCase().endsWith(e.trim().toLowerCase()));
	return /* @__PURE__ */ c("div", {
		onDragOver: (e) => e.preventDefault(),
		onDrop: (e) => {
			e.preventDefault(), t(Array.from(e.dataTransfer.files).filter(r));
		},
		children: n
	});
}
function Je({ recent: e, onOpen: t, onCreate: n, onRecent: r, onRemove: i }) {
	return /* @__PURE__ */ l("main", {
		className: "vo-start-center",
		children: [
			/* @__PURE__ */ l("div", {
				className: "start-head",
				children: [/* @__PURE__ */ c("h1", { children: "VibeOffice" }), /* @__PURE__ */ c(Ke, {})]
			}),
			/* @__PURE__ */ l("section", {
				className: "start-actions",
				children: [/* @__PURE__ */ c("button", {
					onClick: t,
					children: "Open File"
				}), [
					["word", "Quire"],
					["excel", "Ledger"],
					["pptx", "Lectern"]
				].map(([e, t]) => /* @__PURE__ */ l("button", {
					onClick: () => n(e),
					children: [/* @__PURE__ */ c("strong", { children: t }), /* @__PURE__ */ c("span", { children: "Create a file" })]
				}, e))]
			}),
			e.length > 0 && /* @__PURE__ */ l("section", {
				className: "start-recent",
				children: [/* @__PURE__ */ c("h2", { children: "Recent Files" }), e.map((e) => /* @__PURE__ */ l("button", {
					className: "recent-file",
					onClick: () => r(e),
					children: [
						/* @__PURE__ */ c("span", { children: e.name }),
						/* @__PURE__ */ c("small", { children: e.app }),
						/* @__PURE__ */ c("i", {
							onClick: (t) => {
								t.stopPropagation(), i(e);
							},
							children: "×"
						})
					]
				}, `${e.app}:${e.name}:${e.openedAt}`))]
			})
		]
	});
}
function Ye({ script: e = "/sw.js" }) {
	let { updateAvailable: t, applyUpdate: n } = We(e);
	return t ? /* @__PURE__ */ l("div", {
		className: "vo-pwa-update",
		role: "status",
		children: ["A new version is available.", /* @__PURE__ */ c("button", {
			onClick: n,
			children: "Reload"
		})]
	}) : null;
}
function Xe({ visible: e, prompt: t, choices: n, onHide: r, onChoose: i }) {
	return e ? /* @__PURE__ */ l("aside", {
		className: "vo-assistant",
		role: "complementary",
		children: [
			/* @__PURE__ */ c("button", {
				className: "assistant-close",
				onClick: r,
				children: "×"
			}),
			/* @__PURE__ */ c("div", {
				className: "assistant-avatar",
				children: "🐶"
			}),
			t && /* @__PURE__ */ c("p", { children: t }),
			n.map((e) => /* @__PURE__ */ c("button", {
				className: "assistant-choice",
				onClick: () => i(e),
				children: e.label
			}, e.label))
		]
	}) : null;
}
function Ze(e) {
	let t = (e) => e.flatMap((e) => e.sep || !e.label || !e.run ? e.sub ? t(e.sub) : [] : [{
		id: e.id || e.label,
		label: e.label,
		shortcut: e.shortcut,
		keywords: e.keywords,
		disabled: e.disabled,
		run: e.run
	}]);
	return e.flatMap((e) => t(e.items));
}
function Qe({ commands: e, placeholder: t = "Search commands", onRun: n }) {
	let [r, a] = o(""), s = i(() => {
		let t = r.trim().toLowerCase();
		return t ? e.filter((e) => [
			e.id,
			e.label,
			...e.keywords || []
		].join(" ").toLowerCase().includes(t)).slice(0, 12) : [];
	}, [e, r]);
	return /* @__PURE__ */ l("div", {
		className: "cmd-search",
		children: [
			/* @__PURE__ */ c("span", {
				className: "cs-glass",
				children: "⌕"
			}),
			/* @__PURE__ */ c("input", {
				type: "search",
				value: r,
				placeholder: t,
				onChange: (e) => a(e.target.value),
				onKeyDown: (e) => {
					e.key === "Escape" && a("");
				}
			}),
			s.length > 0 && /* @__PURE__ */ c("div", {
				className: "search-results",
				children: s.map((e) => /* @__PURE__ */ l("button", {
					className: "search-hit",
					disabled: e.disabled,
					onMouseDown: (t) => {
						t.preventDefault(), e.disabled || (n?.(e), e.run(), a(""));
					},
					children: [e.label, e.shortcut && /* @__PURE__ */ c("kbd", { children: e.shortcut })]
				}, e.id))
			})
		]
	});
}
function $e({ mode: e, fileName: t, onSubmit: n, onClose: r }) {
	let [i, a] = o("webdav"), [u, d] = o({}), f = (e) => (t) => d((n) => ({
		...n,
		[e]: t.target.value
	}));
	return /* @__PURE__ */ l(Z, {
		title: e === "open" ? "Open from Online" : "Save to Online",
		onClose: r,
		children: [
			/* @__PURE__ */ l("select", {
				value: i,
				onChange: (e) => a(e.target.value),
				children: [
					/* @__PURE__ */ c("option", {
						value: "webdav",
						children: "WebDAV"
					}),
					/* @__PURE__ */ c("option", {
						value: "s3",
						children: "S3 / MinIO"
					}),
					/* @__PURE__ */ c("option", {
						value: "presigned",
						children: "Presigned URL"
					}),
					e === "save" && /* @__PURE__ */ c("option", {
						value: "post",
						children: "HTTP POST"
					})
				]
			}),
			t && /* @__PURE__ */ c("p", { children: t }),
			/* @__PURE__ */ l("label", { children: ["URL / Endpoint", /* @__PURE__ */ c("input", {
				autoFocus: !0,
				value: u.url || u.endpoint || "",
				onChange: f(i === "s3" ? "endpoint" : "url"),
				placeholder: "https://…"
			})] }),
			i === "webdav" && /* @__PURE__ */ l(s, { children: [/* @__PURE__ */ l("label", { children: ["Username", /* @__PURE__ */ c("input", {
				value: u.username || "",
				onChange: f("username")
			})] }), /* @__PURE__ */ l("label", { children: ["Password", /* @__PURE__ */ c("input", {
				type: "password",
				value: u.password || "",
				onChange: f("password")
			})] })] }),
			i === "s3" && /* @__PURE__ */ l(s, { children: [
				/* @__PURE__ */ l("label", { children: ["Bucket", /* @__PURE__ */ c("input", {
					value: u.bucket || "",
					onChange: f("bucket")
				})] }),
				/* @__PURE__ */ l("label", { children: ["Object key", /* @__PURE__ */ c("input", {
					value: u.key || "",
					onChange: f("key")
				})] }),
				/* @__PURE__ */ l("label", { children: ["Region", /* @__PURE__ */ c("input", {
					value: u.region || "",
					onChange: f("region")
				})] }),
				/* @__PURE__ */ l("label", { children: ["Access key", /* @__PURE__ */ c("input", {
					value: u.accessKeyId || "",
					onChange: f("accessKeyId")
				})] }),
				/* @__PURE__ */ l("label", { children: ["Secret key", /* @__PURE__ */ c("input", {
					type: "password",
					value: u.secretAccessKey || "",
					onChange: f("secretAccessKey")
				})] })
			] }),
			i === "presigned" && e === "save" && /* @__PURE__ */ l("label", { children: ["PUT URL", /* @__PURE__ */ c("input", {
				value: u.putUrl || "",
				onChange: f("putUrl")
			})] }),
			i === "post" && /* @__PURE__ */ l("label", { children: ["Multipart field", /* @__PURE__ */ c("input", {
				value: u.field || "file",
				onChange: f("field")
			})] }),
			/* @__PURE__ */ l("div", { children: [/* @__PURE__ */ c("button", {
				onClick: r,
				children: "Cancel"
			}), /* @__PURE__ */ c("button", {
				onClick: () => n({
					kind: i,
					...u
				}),
				disabled: !u.url && !u.endpoint,
				children: e === "open" ? "Open" : "Save"
			})] })
		]
	});
}
// src/index.ts
function $(e) {
	try {
		return decodeURIComponent(new URL(e).pathname.split("/").filter(Boolean).pop() || "document");
	} catch {
		return "document";
	}
}
async function et(e) {
	if (e.kind === "s3") return ce(e);
	if (e.kind === "webdav") return ue(e);
	if (e.kind === "http") return fe(e);
	let t = await fetch(e.url);
	if (!t.ok) throw Error(`Download failed: ${t.status} ${t.statusText}`);
	return {
		blob: await t.blob(),
		name: $(e.url)
	};
}
async function tt(e, t, n) {
	if (t = await t, e.kind === "post") return pe(e, t, n);
	if (e.kind === "s3") return le(e, t);
	if (e.kind === "webdav") return de(e, t);
	if (!e.putUrl) throw Error("The online source is read-only");
	let r = await fetch(e.putUrl, {
		method: "PUT",
		headers: { "Content-Type": t.type || "application/octet-stream" },
		body: t
	});
	if (!r.ok) throw Error(`Upload failed: ${r.status} ${r.statusText}`);
}
function nt(e, t) {
	let n = document.createElement("a");
	n.href = URL.createObjectURL(e), n.download = t, n.click(), setTimeout(() => URL.revokeObjectURL(n.href), 0);
}
var rt = class extends e {
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
export { Xe as Assistant, Qe as CommandSearch, Re as ContextMenu, Z as Dialog, b as El, rt as ErrorBoundary, qe as FileDropZone, Ke as LanguageSelect, Me as MenuBar, G as NS, Le as OnlineDialog, $e as OnlineSourceDialog, Oe as Package, Ye as PwaUpdate, Fe as Ribbon, Je as StartCenter, Ie as StatusBar, Pe as Toolbar, y as Txt, ne as attrs, ie as chars, Ze as commandsFromMenus, ke as contentTypeForExt, N as crc32, Ee as decode, C as decodeBytes, nt as downloadBlob, Te as encode, h as esc, v as escText, et as fetchOnline, K as formats, z as inflateRaw, L as inflateRawJS, oe as mkEl, $ as nameFromUrl, Ae as newPackage, S as parseXml, U as readZip, Y as resolveTarget, tt as saveOnline, ae as scanXml, w as serializeXml, m as unescape, ee as unx, Ge as useAssistant, He as useDirtyGuard, Ue as useI18n, Ve as useLook, We as usePwa, Be as useRecentFiles, De as uuid, W as writeZip, je as xmlPart };
