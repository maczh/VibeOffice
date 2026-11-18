import { c as e, f as t, g as n, p as r, s as i } from "./engine-BYKPaHub.js";
import { forwardRef as a, useEffect as o, useImperativeHandle as s, useRef as c, useState as l } from "react";
import { ErrorBoundary as u, downloadBlob as d, fetchOnline as f, saveOnline as p } from "@voffice/react-shared";
import { jsx as m, jsxs as h } from "react/jsx-runtime";
// src/index.tsx
function g(e, t = "docx") {
	let i = n();
	return i.format = t, i.blocks = e.split(/\r?\n/).map((e) => ({
		...r(),
		runs: e ? [{ text: e }] : []
	})), i;
}
var _ = a(function({ src: r, fileName: a = "Document.docx", readonly: _, onReady: v, onChange: y, onSave: b, onError: x }, S) {
	let [C, w] = l(a), [T, E] = l(!1), [D, O] = l(""), [k, A] = l(""), j = c(n()), M = c(null), N = c("word/document.xml"), P = c(null), F = (e) => {
		E(e), y?.(e);
	}, I = async (n, r) => {
		try {
			O("Opening…");
			let i = n instanceof Blob ? {
				blob: n,
				name: r ?? (n instanceof File ? n.name : "Document.docx")
			} : await f(typeof n == "string" ? {
				kind: "http",
				url: n
			} : n), a = await e(new Uint8Array(await i.blob.arrayBuffer()));
			j.current = a.doc, M.current = a.pkg, N.current = a.part, P.current = n, A(t(a.doc)), w(i.name), F(!1), v?.();
		} catch (e) {
			x?.(e instanceof Error ? e : Error(String(e)));
		} finally {
			O("");
		}
	}, L = async () => i(j.current, {
		pkg: M.current ?? void 0,
		part: N.current,
		format: j.current.format
	}), R = async (e = C) => {
		let t = await L(), n = /\.[a-z0-9]+$/i.test(e) ? e : `${e}.docx`;
		return d(t, n), F(!1), t;
	}, z = async () => {
		let e = await L(), t = P.current;
		t && typeof t != "string" && "kind" in t && t.kind !== "http" ? await p(t, e, C) : b?.(e, C), F(!1);
	}, B = () => {
		j.current = n(), M.current = null, P.current = null, A(""), w("Document.docx"), F(!1);
	};
	s(S, () => ({
		open: I,
		newDocument: B,
		getBlob: L,
		save: z,
		saveAs: R,
		saveTo: async (e, t) => p(e, await L(), t ?? C),
		get name() {
			return C;
		},
		get dirty() {
			return T;
		}
	}), [
		C,
		T,
		k
	]), o(() => {
		r && I(r);
	}, []);
	let V = (e) => {
		A(e), j.current = g(e, j.current.format), F(!0);
	};
	return /* @__PURE__ */ m(u, { children: /* @__PURE__ */ h("div", {
		className: "vo-office",
		children: [
			/* @__PURE__ */ h("div", {
				className: "vo-title",
				children: [
					"Quire — ",
					C,
					T ? " *" : ""
				]
			}),
			/* @__PURE__ */ h("div", {
				className: "vo-toolbar",
				children: [
					/* @__PURE__ */ m("button", {
						onClick: () => void z(),
						children: "Save"
					}),
					/* @__PURE__ */ m("button", {
						onClick: () => void R(),
						children: "Save As"
					}),
					/* @__PURE__ */ m("button", {
						onClick: B,
						children: "New"
					}),
					D && /* @__PURE__ */ m("span", { children: D })
				]
			}),
			/* @__PURE__ */ m("main", {
				className: "vo-editor",
				children: /* @__PURE__ */ m("textarea", {
					readOnly: _,
					value: k,
					onChange: (e) => V(e.target.value),
					"aria-label": "Document editor"
				})
			}),
			/* @__PURE__ */ h("div", {
				className: "vo-status",
				children: [k.trim().split(/\s+/).filter(Boolean).length, " words"]
			})
		]
	}) });
});
export { _ as RWord, _ as default };
