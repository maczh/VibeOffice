import { a as e, d as t, i as n, l as r } from "./engine-B7fFY5dR.js";
import { forwardRef as i, useEffect as a, useImperativeHandle as o, useRef as s, useState as c } from "react";
import { ErrorBoundary as l, downloadBlob as u, fetchOnline as d, saveOnline as f } from "@voffice/react-shared";
import { jsx as p, jsxs as m } from "react/jsx-runtime";
// src/index.tsx
function h(e) {
	return e.slides.map((e, t) => `Slide ${t + 1}\n${e.shapes.flatMap((e) => e.text ?? []).flatMap((e) => e.runs.map((e) => e.text)).join("")}`).join("\n\n");
}
function g(e) {
	let n = r();
	return n.slides[0].shapes = [t(914400, 914400, 10058400, 4572e3, e)], n;
}
var _ = i(function({ src: t, fileName: i = "Presentation.pptx", readonly: _, onReady: v, onChange: y, onSave: b, onError: x }, S) {
	let [C, w] = c(i), [T, E] = c(!1), [D, O] = c(""), [k, A] = c(""), j = s(r()), M = s(null), N = s([]), P = s(null), F = (e) => {
		E(e), y?.(e);
	}, I = async (t, n) => {
		try {
			O("Opening…");
			let r = t instanceof Blob ? {
				blob: t,
				name: n ?? (t instanceof File ? t.name : "Presentation.pptx")
			} : await d(typeof t == "string" ? {
				kind: "http",
				url: t
			} : t), i = await e(new Uint8Array(await r.blob.arrayBuffer()));
			j.current = i.pres, M.current = i.pkg, N.current = i.slideParts, P.current = t, A(h(i.pres)), w(r.name), F(!1), v?.();
		} catch (e) {
			x?.(e instanceof Error ? e : Error(String(e)));
		} finally {
			O("");
		}
	}, L = async () => n(j.current, {
		pkg: M.current ?? void 0,
		slideParts: N.current,
		format: j.current.format
	}), R = async (e = C) => {
		let t = await L(), n = /\.[a-z0-9]+$/i.test(e) ? e : `${e}.pptx`;
		return u(t, n), F(!1), t;
	}, z = async () => {
		let e = await L(), t = P.current;
		t && typeof t != "string" && "kind" in t && t.kind !== "http" ? await f(t, e, C) : b?.(e, C), F(!1);
	}, B = () => {
		j.current = r(), M.current = null, N.current = [], P.current = null, A(""), w("Presentation.pptx"), F(!1);
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
		t && I(t);
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
					"Lectern — ",
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
					"aria-label": "Presentation editor"
				})
			}),
			/* @__PURE__ */ p("div", {
				className: "vo-status",
				children: "Ready"
			})
		]
	}) });
});
export { _ as RPptx, _ as default };
