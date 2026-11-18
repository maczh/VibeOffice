import { C as e, f as t, p as n, x as r } from "./engine-GRzmteNh.js";
import { forwardRef as i, useEffect as a, useImperativeHandle as o, useRef as s, useState as c } from "react";
import { ErrorBoundary as l, downloadBlob as u, fetchOnline as d, saveOnline as f } from "@voffice/react-shared";
import { jsx as p, jsxs as m } from "react/jsx-runtime";
// src/index.tsx
function h(e) {
	return [...e.sheets[e.active].cells.entries()].sort(([e], [t]) => e.localeCompare(t)).map(([e, t]) => `${e}\t${t.f ? "=" + t.f : String(t.v ?? "")}`).join("\n");
}
function g(t) {
	let n = r(), i = n.sheets[0];
	return t.split(/\r?\n/).forEach((t, n) => t.split("	").forEach((t, r) => {
		if (t.startsWith("=")) e(i, n, r, { f: t.slice(1) });
		else if (t !== "") {
			let a = Number(t);
			e(i, n, r, Number.isFinite(a) && t.trim() !== "" ? {
				v: a,
				t: "n"
			} : {
				v: t,
				t: "s"
			});
		}
	})), n;
}
var _ = i(function({ src: e, fileName: i = "Book1.xlsx", readonly: _, onReady: v, onChange: y, onSave: b, onError: x }, S) {
	let [C, w] = c(i), [T, E] = c(!1), [D, O] = c(""), [k, A] = c(""), j = s(r()), M = s(null), N = s([]), P = s(null), F = (e) => {
		E(e), y?.(e);
	}, I = async (e, t) => {
		try {
			O("Opening…");
			let r = e instanceof Blob ? {
				blob: e,
				name: t ?? (e instanceof File ? e.name : "Book1.xlsx")
			} : await d(typeof e == "string" ? {
				kind: "http",
				url: e
			} : e), i = await n(new Uint8Array(await r.blob.arrayBuffer()));
			j.current = i.wb, M.current = i.pkg, N.current = i.sheetParts, P.current = e, A(h(i.wb)), w(r.name), F(!1), v?.();
		} catch (e) {
			x?.(e instanceof Error ? e : Error(String(e)));
		} finally {
			O("");
		}
	}, L = async () => t(j.current, {
		pkg: M.current ?? void 0,
		sheetParts: N.current,
		format: "xlsx"
	}), R = async (e = C) => {
		let t = await L(), n = /\.[a-z0-9]+$/i.test(e) ? e : `${e}.xlsx`;
		return u(t, n), F(!1), t;
	}, z = async () => {
		let e = await L(), t = P.current;
		t && typeof t != "string" && "kind" in t && t.kind !== "http" ? await f(t, e, C) : b?.(e, C), F(!1);
	}, B = () => {
		j.current = r(), M.current = null, N.current = [], P.current = null, A(""), w("Book1.xlsx"), F(!1);
	};
	o(S, () => ({
		open: I,
		newDocument: B,
		getBlob: L,
		save: z,
		saveAs: R,
		saveTo: async (e, t) => f(e, await L(), t ?? C),
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
	]), a(() => {
		e && I(e);
	}, []);
	let V = (e) => {
		A(e), j.current = g(e), F(!0);
	};
	return /* @__PURE__ */ p(l, { children: /* @__PURE__ */ m("div", {
		className: "vo-office",
		children: [
			/* @__PURE__ */ m("div", {
				className: "vo-title",
				children: [
					"Ledger — ",
					C,
					T ? " *" : ""
				]
			}),
			/* @__PURE__ */ m("div", {
				className: "vo-toolbar",
				children: [
					/* @__PURE__ */ p("button", {
						onClick: () => void z(),
						children: "Save"
					}),
					/* @__PURE__ */ p("button", {
						onClick: () => void R(),
						children: "Save As"
					}),
					/* @__PURE__ */ p("button", {
						onClick: B,
						children: "New"
					}),
					D && /* @__PURE__ */ p("span", { children: D })
				]
			}),
			/* @__PURE__ */ p("main", {
				className: "vo-editor",
				children: /* @__PURE__ */ p("textarea", {
					readOnly: _,
					value: k,
					onChange: (e) => V(e.target.value),
					"aria-label": "Spreadsheet editor"
				})
			}),
			/* @__PURE__ */ p("div", {
				className: "vo-status",
				children: "Ready"
			})
		]
	}) });
});
export { _ as RExcel, _ as default };
