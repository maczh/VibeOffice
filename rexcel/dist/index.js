import { C as e, b as t, f as n, l as r, p as i, x as a } from "./engine-CYvYT4sX.js";
import { forwardRef as o, useEffect as s, useImperativeHandle as c, useRef as l, useState as u } from "react";
import { ContextMenu as d, Dialog as ee, ErrorBoundary as f, MenuBar as p, OnlineDialog as m, Ribbon as h, StatusBar as g, Toolbar as _, downloadBlob as v, fetchOnline as te, saveOnline as y, useDirtyGuard as ne } from "@voffice/react-shared";
import { jsx as b, jsxs as x } from "react/jsx-runtime";
// src/index.tsx
function re(e) {
	return [...e.sheets[e.active].cells.entries()].sort(([e], [t]) => e.localeCompare(t)).map(([e, t]) => `${e}\t${t.f ? "=" + t.f : String(t.v ?? "")}`).join("\n");
}
function S(t) {
	let n = a(), i = n.sheets[0];
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
	})), r(n), n;
}
var C = o(function({ src: e, fileName: r = "Book1.xlsx", readonly: o, onReady: C, onChange: ie, onSave: w, onError: T }, ae) {
	let [E, D] = u(r), [O, k] = u(!1), [A, j] = u(""), [M, N] = u(""), [P, F] = u(null), [I, L] = u(null), R = l(a()), z = l(null), B = l([]), V = l(null), H = l([]), U = l([]), W = (e) => {
		k(e), ie?.(e);
	};
	ne(O);
	let G = async (e, t) => {
		try {
			j("Opening…");
			let n = e instanceof Blob ? {
				blob: e,
				name: t ?? (e instanceof File ? e.name : "Book1.xlsx")
			} : await te(typeof e == "string" ? {
				kind: "http",
				url: e
			} : e), r = await i(new Uint8Array(await n.blob.arrayBuffer()));
			R.current = r.wb, z.current = r.pkg, B.current = r.sheetParts, V.current = e, N(re(r.wb)), D(n.name), W(!1), C?.();
		} catch (e) {
			T?.(e instanceof Error ? e : Error(String(e)));
		} finally {
			j("");
		}
	}, K = async () => n(R.current, {
		pkg: z.current ?? void 0,
		sheetParts: B.current,
		format: "xlsx"
	}), q = async (e = E) => {
		let t = await K(), n = /\.[a-z0-9]+$/i.test(e) ? e : `${e}.xlsx`;
		return v(t, n), W(!1), t;
	}, J = async () => {
		let e = await K(), t = V.current;
		t && typeof t != "string" && "kind" in t && t.kind !== "http" ? await y(t, e, E) : w?.(e, E), W(!1);
	}, Y = () => {
		R.current = a(), z.current = null, B.current = [], V.current = null, H.current = [], U.current = [], N(""), D("Book1.xlsx"), W(!1);
	};
	c(ae, () => ({
		open: G,
		newDocument: Y,
		getBlob: K,
		save: J,
		saveAs: q,
		saveTo: async (e, t) => y(e, await K(), t ?? E),
		get name() {
			return E;
		},
		get dirty() {
			return O;
		}
	}), [
		E,
		O,
		M
	]), s(() => {
		e && G(e);
	}, []), s(() => {
		let e = (e) => {
			if (!(e.ctrlKey || e.metaKey)) return;
			let t = e.key.toLowerCase();
			t === "s" ? (e.preventDefault(), J()) : t === "z" ? (e.preventDefault(), e.shiftKey ? Q() : Z()) : t === "y" ? (e.preventDefault(), Q()) : t === "n" ? (e.preventDefault(), Y()) : t === "o" && (e.preventDefault(), document.querySelector("#rexcel-file")?.click());
		};
		return window.addEventListener("keydown", e), () => window.removeEventListener("keydown", e);
	}, [
		M,
		E,
		O
	]);
	let X = (e) => {
		e !== M && (H.current.push(M), U.current = [], H.current.length > 50 && H.current.shift()), N(e), R.current = S(e), W(!0);
	}, Z = () => {
		let e = H.current.pop();
		e !== void 0 && (U.current.push(M), N(e), R.current = S(e), W(!0));
	}, Q = () => {
		let e = U.current.pop();
		e !== void 0 && (H.current.push(M), N(e), R.current = S(e), W(!0));
	}, oe = () => {
		let e = t(`Sheet${R.current.sheets.length + 1}`);
		R.current.sheets.push(e), W(!0);
	}, se = () => X(`\n${M}`), ce = () => X(M.split(/\r?\n/).map((e) => `\t${e}`).join("\n")), le = [
		{
			label: "File",
			items: [
				{
					label: "New",
					run: Y
				},
				{
					label: "Open…",
					run: () => document.querySelector("#rexcel-file")?.click()
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
				},
				{
					label: "Print…",
					run: () => window.print()
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
					label: "Clear Contents",
					run: () => X("")
				}
			]
		},
		{
			label: "View",
			items: [
				{
					label: "Gridlines",
					checked: !0
				},
				{ label: "Zoom In" },
				{ label: "Zoom Out" }
			]
		},
		{
			label: "Insert",
			items: [
				{ label: "Rows" },
				{ label: "Columns" },
				{ label: "Sheet" }
			]
		},
		{
			label: "Format",
			items: [
				{ label: "Bold" },
				{ label: "Italic" },
				{ label: "Currency" }
			]
		},
		{
			label: "Data",
			items: [{ label: "Sort Ascending" }, { label: "Sort Descending" }]
		},
		{
			label: "Help",
			items: [{ label: "About Ledger" }]
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
			run: () => document.querySelector("#rexcel-file")?.click()
		},
		{
			id: "online-open",
			label: "Online Open",
			run: () => L("open")
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
			run: () => L("save")
		},
		{
			id: "about",
			label: "About",
			run: () => L("about")
		},
		{
			id: "rows",
			label: "Insert Rows",
			run: se
		},
		{
			id: "cols",
			label: "Insert Columns",
			run: ce
		},
		{
			id: "sheet",
			label: "New Sheet",
			run: oe
		},
		{
			id: "print",
			label: "Print",
			run: () => window.print()
		}
	];
	return /* @__PURE__ */ b(f, { children: /* @__PURE__ */ x("div", {
		className: "vo-office",
		children: [
			/* @__PURE__ */ x("div", {
				className: "vo-title",
				children: [
					"Ledger — ",
					E,
					O ? " *" : ""
				]
			}),
			/* @__PURE__ */ b("input", {
				id: "rexcel-file",
				hidden: !0,
				type: "file",
				accept: ".xlsx,.xlsm,.xltx,.xltm,.csv,.txt",
				onChange: (e) => {
					let t = e.target.files?.[0];
					t && G(t);
				}
			}),
			/* @__PURE__ */ b(p, { menus: le }),
			/* @__PURE__ */ b(_, { items: $ }),
			/* @__PURE__ */ b(h, { tabs: [
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
							id: "rows",
							label: "Rows"
						},
						{
							id: "cols",
							label: "Columns"
						},
						{
							id: "sheet",
							label: "Sheet"
						}
					]
				},
				{
					id: "data",
					label: "Data",
					items: [{
						id: "asc",
						label: "A→Z"
					}, {
						id: "desc",
						label: "Z→A"
					}]
				}
			] }),
			/* @__PURE__ */ b("main", {
				className: "vo-editor",
				children: /* @__PURE__ */ b("textarea", {
					readOnly: o,
					value: M,
					onContextMenu: (e) => {
						e.preventDefault(), F({
							x: e.clientX,
							y: e.clientY
						});
					},
					onChange: (e) => X(e.target.value),
					"aria-label": "Spreadsheet editor"
				})
			}),
			/* @__PURE__ */ b(g, { children: A || "Ready" }),
			P && /* @__PURE__ */ b(d, {
				x: P.x,
				y: P.y,
				onClose: () => F(null),
				items: [
					{
						label: "Cut",
						run: () => {
							document.execCommand("cut");
						}
					},
					{
						label: "Copy",
						run: () => {
							document.execCommand("copy");
						}
					},
					{
						label: "Paste",
						run: () => {
							document.execCommand("paste");
						}
					},
					{
						sep: !0,
						label: ""
					},
					{
						label: "Clear Contents",
						run: () => X("")
					},
					{
						label: "Insert Rows",
						disabled: !0
					},
					{
						label: "Insert Columns",
						disabled: !0
					}
				]
			}),
			I === "open" && /* @__PURE__ */ b(m, {
				mode: "open",
				onClose: () => L(null),
				onSubmit: (e) => {
					L(null), G(e);
				}
			}),
			I === "save" && /* @__PURE__ */ b(m, {
				mode: "save",
				fileName: E,
				onClose: () => L(null),
				onSubmit: (e) => {
					L(null), y({
						kind: "presigned",
						url: e,
						putUrl: e
					}, K(), E);
				}
			}),
			I === "about" && /* @__PURE__ */ b(ee, {
				title: "About Ledger",
				onClose: () => L(null),
				children: /* @__PURE__ */ b("p", { children: "Ledger is the VibeOffice spreadsheet editor." })
			})
		]
	}) });
});
export { C as RExcel, C as default };
