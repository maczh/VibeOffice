import { a as e, d as t, i as n, l as r, u as i } from "./engine-agthuCk9.js";
import { forwardRef as a, useEffect as o, useImperativeHandle as ee, useRef as s, useState as c } from "react";
import { ContextMenu as l, Dialog as u, ErrorBoundary as d, MenuBar as f, OnlineDialog as p, Ribbon as m, StatusBar as h, Toolbar as te, downloadBlob as ne, fetchOnline as g, saveOnline as _, useDirtyGuard as v } from "@voffice/react-shared";
import { jsx as y, jsxs as b } from "react/jsx-runtime";
// src/index.tsx
function x(e) {
	return e.slides.map((e, t) => `Slide ${t + 1}\n${e.shapes.flatMap((e) => e.text ?? []).flatMap((e) => e.runs.map((e) => e.text)).join("")}`).join("\n\n");
}
function S(e) {
	let n = r();
	return n.slides = e.split(/\n\s*\n/).map((e, r) => {
		let a = r === 0 ? n.slides[0] : i();
		return a.shapes = [t(914400, 914400, 10058400, 4572e3, e.replace(/^Slide \d+\n/, ""))], a;
	}), n;
}
var C = a(function({ src: t, fileName: i = "Presentation.pptx", readonly: a, onReady: C, onChange: re, onSave: ie, onError: ae }, oe) {
	let [w, T] = c(i), [E, D] = c(!1), [O, k] = c(""), [A, j] = c(""), [M, N] = c(null), [P, F] = c(null), [I, L] = c(!1), R = s(r()), z = s(null), B = s([]), V = s(null), H = s([]), U = s([]), W = (e) => {
		D(e), re?.(e);
	};
	v(E);
	let G = async (t, n) => {
		try {
			k("Opening…");
			let r = t instanceof Blob ? {
				blob: t,
				name: n ?? (t instanceof File ? t.name : "Presentation.pptx")
			} : await g(typeof t == "string" ? {
				kind: "http",
				url: t
			} : t), i = await e(new Uint8Array(await r.blob.arrayBuffer()));
			R.current = i.pres, z.current = i.pkg, B.current = i.slideParts, V.current = t, j(x(i.pres)), T(r.name), W(!1), C?.();
		} catch (e) {
			ae?.(e instanceof Error ? e : Error(String(e)));
		} finally {
			k("");
		}
	}, K = async () => n(R.current, {
		pkg: z.current ?? void 0,
		slideParts: B.current,
		format: R.current.format
	}), q = async (e = w) => {
		let t = await K(), n = /\.[a-z0-9]+$/i.test(e) ? e : `${e}.pptx`;
		return ne(t, n), W(!1), t;
	}, J = async () => {
		let e = await K(), t = V.current;
		t && typeof t != "string" && "kind" in t && t.kind !== "http" ? await _(t, e, w) : ie?.(e, w), W(!1);
	}, Y = () => {
		R.current = r(), z.current = null, B.current = [], V.current = null, H.current = [], U.current = [], j(""), T("Presentation.pptx"), W(!1);
	};
	ee(oe, () => ({
		open: G,
		newDocument: Y,
		getBlob: K,
		save: J,
		saveAs: q,
		saveTo: async (e, t) => _(e, await K(), t ?? w),
		get name() {
			return w;
		},
		get dirty() {
			return E;
		}
	}), [
		w,
		E,
		A
	]), o(() => {
		t && G(t);
	}, []), o(() => {
		if (!I) return;
		let e = (e) => {
			e.key === "Escape" && L(!1);
		};
		return window.addEventListener("keydown", e), () => window.removeEventListener("keydown", e);
	}, [I]), o(() => {
		let e = (e) => {
			if (!(e.ctrlKey || e.metaKey)) return;
			let t = e.key.toLowerCase();
			t === "s" ? (e.preventDefault(), J()) : t === "z" ? (e.preventDefault(), e.shiftKey ? Q() : Z()) : t === "y" ? (e.preventDefault(), Q()) : t === "n" ? (e.preventDefault(), Y()) : t === "o" && (e.preventDefault(), document.querySelector("#rpptx-file")?.click());
		};
		return window.addEventListener("keydown", e), () => window.removeEventListener("keydown", e);
	}, [
		A,
		w,
		E
	]);
	let X = (e) => {
		e !== A && (H.current.push(A), U.current = [], H.current.length > 50 && H.current.shift()), j(e), R.current = S(e), W(!0);
	}, Z = () => {
		let e = H.current.pop();
		e !== void 0 && (U.current.push(A), j(e), R.current = S(e), W(!0));
	}, Q = () => {
		let e = U.current.pop();
		e !== void 0 && (H.current.push(A), j(e), R.current = S(e), W(!0));
	}, se = [
		{
			label: "File",
			items: [
				{
					label: "New",
					run: Y
				},
				{
					label: "Open…",
					run: () => document.querySelector("#rpptx-file")?.click()
				},
				{
					sep: !0,
					label: ""
				},
				{
					label: "Save",
					run: J
				},
				{
					label: "Save As…",
					run: async () => {
						await q();
					}
				}
			]
		},
		{
			label: "Edit",
			items: [
				{
					label: "Undo",
					disabled: !0
				},
				{
					label: "Redo",
					disabled: !0
				},
				{
					sep: !0,
					label: ""
				},
				{
					label: "Delete Shape",
					disabled: !0
				}
			]
		},
		{
			label: "View",
			items: [
				{ label: "Fit to Window" },
				{ label: "Zoom In" },
				{ label: "Slide Show" }
			]
		},
		{
			label: "Insert",
			items: [
				{ label: "New Slide" },
				{ label: "Duplicate Slide" },
				{ label: "Text Box" },
				{ label: "Picture…" },
				{ label: "Shape" }
			]
		},
		{
			label: "Format",
			items: [
				{ label: "Bold" },
				{ label: "Italic" },
				{ label: "Shape Fill" },
				{ label: "Slide Background…" }
			]
		},
		{
			label: "Slide Show",
			items: [{ label: "From Beginning" }, { label: "From Current Slide" }]
		},
		{
			label: "Help",
			items: [{ label: "About Lectern" }]
		}
	], $ = [
		{
			id: "new",
			label: "New",
			run: Y
		},
		{
			id: "open",
			label: "Open",
			run: () => document.querySelector("#rpptx-file")?.click()
		},
		{
			id: "online-open",
			label: "Online Open",
			run: () => F("open")
		},
		{
			id: "save",
			label: "Save",
			run: J
		},
		{
			id: "undo",
			label: "Undo",
			run: Z
		},
		{
			id: "redo",
			label: "Redo",
			run: Q
		},
		{
			id: "online-save",
			label: "Online Save",
			run: () => F("save")
		},
		{
			id: "about",
			label: "About",
			run: () => F("about")
		},
		{
			id: "slide",
			label: "New Slide",
			run: () => {
				X(`${A}\n\nSlide ${R.current.slides.length + 1}\n`), W(!0);
			}
		},
		{
			id: "show",
			label: "Slide Show",
			run: () => L(!0)
		}
	];
	return /* @__PURE__ */ y(d, { children: /* @__PURE__ */ b("div", {
		className: "vo-office",
		children: [
			/* @__PURE__ */ b("div", {
				className: "vo-title",
				children: [
					"Lectern — ",
					w,
					E ? " *" : ""
				]
			}),
			/* @__PURE__ */ y("input", {
				id: "rpptx-file",
				hidden: !0,
				type: "file",
				accept: ".pptx,.pptm,.ppsx,.ppsm,.potx,.potm",
				onChange: (e) => {
					let t = e.target.files?.[0];
					t && G(t);
				}
			}),
			/* @__PURE__ */ y(f, { menus: se }),
			/* @__PURE__ */ y(te, { items: $ }),
			/* @__PURE__ */ y(m, { tabs: [
				{
					id: "home",
					label: "Home",
					items: $
				},
				{
					id: "insert",
					label: "Insert",
					items: [
						{
							id: "slide",
							label: "New Slide"
						},
						{
							id: "picture",
							label: "Picture"
						},
						{
							id: "shape",
							label: "Shapes"
						},
						{
							id: "text",
							label: "Text Box"
						}
					]
				},
				{
					id: "show",
					label: "Slide Show",
					items: [{
						id: "start",
						label: "From Start"
					}, {
						id: "current",
						label: "Current"
					}]
				}
			] }),
			/* @__PURE__ */ y("main", {
				className: "vo-editor",
				children: /* @__PURE__ */ y("textarea", {
					readOnly: a,
					value: A,
					onContextMenu: (e) => {
						e.preventDefault(), N({
							x: e.clientX,
							y: e.clientY
						});
					},
					onChange: (e) => X(e.target.value),
					"aria-label": "Presentation editor"
				})
			}),
			/* @__PURE__ */ y(h, { children: O || "Ready" }),
			M && /* @__PURE__ */ y(l, {
				x: M.x,
				y: M.y,
				onClose: () => N(null),
				items: [
					{
						label: "Cut",
						disabled: !0
					},
					{
						label: "Copy",
						disabled: !0
					},
					{
						label: "Paste",
						disabled: !0
					},
					{
						sep: !0,
						label: ""
					},
					{
						label: "Delete Shape",
						disabled: !0
					},
					{
						label: "New Slide",
						disabled: !0
					}
				]
			}),
			P === "open" && /* @__PURE__ */ y(p, {
				mode: "open",
				onClose: () => F(null),
				onSubmit: (e) => {
					F(null), G(e);
				}
			}),
			P === "save" && /* @__PURE__ */ y(p, {
				mode: "save",
				fileName: w,
				onClose: () => F(null),
				onSubmit: (e) => {
					F(null), _({
						kind: "presigned",
						url: e,
						putUrl: e
					}, K(), w);
				}
			}),
			P === "about" && /* @__PURE__ */ y(u, {
				title: "About Lectern",
				onClose: () => F(null),
				children: /* @__PURE__ */ y("p", { children: "Lectern is the VibeOffice presentation editor." })
			}),
			I && /* @__PURE__ */ b("div", {
				className: "vo-show",
				role: "dialog",
				"aria-label": "Slide show",
				onClick: () => L(!1),
				children: [/* @__PURE__ */ y("pre", { children: x(R.current) }), /* @__PURE__ */ y("span", { children: "Click or press Escape to exit" })]
			})
		]
	}) });
});
export { C as RPptx, C as default };
