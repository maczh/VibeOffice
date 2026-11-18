import { $ as e, F as t, G as n, K as r, U as i, V as a, W as o, at as s, c, ct as l, et as ee, it as u } from "./engine-D-DARJG-.js";
import { forwardRef as d, useEffect as f, useImperativeHandle as te, useMemo as p, useRef as m, useState as h } from "react";
import { ContextMenu as ne, Dialog as g, ErrorBoundary as re, MenuBar as ie, OnlineDialog as ae, Ribbon as oe, StatusBar as se, Toolbar as _, downloadBlob as ce, fetchOnline as le, saveOnline as v, useDirtyGuard as ue } from "@voffice/react-shared";
import { Fragment as y, jsx as b, jsxs as x } from "react/jsx-runtime";
// src/find.ts
var S = (e) => e.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), C = (e) => {
	let t = e.toUpperCase().replace(/[^A-Z]/g, "");
	if (!t) return "";
	let n = {
		B: 1,
		F: 1,
		P: 1,
		V: 1,
		C: 2,
		G: 2,
		J: 2,
		K: 2,
		Q: 2,
		S: 2,
		X: 2,
		Z: 2,
		D: 3,
		T: 3,
		L: 4,
		M: 5,
		N: 5,
		R: 6
	}, r = t[0], i = n[t[0]] || 0;
	for (let e of t.slice(1)) {
		let t = n[e] || 0;
		t && t !== i && (r += t), e !== "H" && e !== "W" && (i = t);
	}
	return (r + "000").slice(0, 4);
};
function w(e) {
	let t = e.wildcards ? e.text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\\\*/g, "[^\\r]*?").replace(/\\\?/g, "[^\\r]") : S(e.text);
	return e.allForms && (t = `${S(e.text.replace(/(ing|ed|es|s)$/i, ""))}(?:s|es|ed|d|ing|er|ers|ly|'s)?`), e.wholeWord && (t = `(?<![\\p{L}\\p{N}_])${t}(?![\\p{L}\\p{N}_])`), t;
}
function T(e, t) {
	if (!t.text) return [];
	let n;
	try {
		n = new RegExp(w(t), `gu${t.matchCase || t.wildcards ? "" : "i"}`);
	} catch {
		throw Error("The Find What text contains an invalid pattern.");
	}
	let r = t.soundsLike ? C(t.text) : "", i = [], a;
	for (; a = n.exec(e);) r && C(a[0]) !== r || i.push({
		index: a.index,
		end: a.index + a[0].length,
		value: a[0],
		groups: a.slice(1)
	}), a[0] === "" && n.lastIndex++;
	return i;
}
function de(e, t) {
	let n = T(e, t), r = e;
	for (let e = n.length - 1; e >= 0; e--) {
		let i = n[e], a = (t.replace || "").replace(/\\(\d+)/g, (e, t) => i.groups[Number(t) - 1] || "").replace(/\^&/g, i.value).replace(/\^p/g, "\n").replace(/\^t/g, "	");
		r = r.slice(0, i.index) + a + r.slice(i.end);
	}
	return {
		value: r,
		count: n.length
	};
}
// src/quire-manifest.ts
var E = {
	migratedServices: [
		"find.js",
		"markdown.js",
		"mdio.js",
		"htmlio.js",
		"autocorrect.js",
		"mailmerge.js",
		"tables.js",
		"fields.js",
		"spell.js",
		"drawing.js",
		"drawtool.js",
		"rtf.js",
		"review.js",
		"preserve.js",
		"preserve-drawing.js"
	],
	migratedUi: [
		"app.js",
		"dialogs.js",
		"dialogs2.js",
		"dialogs3.js",
		"ribbon.js",
		"rulers.js",
		"panes.js",
		"editor.js",
		"outline.js"
	],
	engineModules: [
		"commands.js",
		"dmodel.js",
		"docx-read.js",
		"docx-write.js",
		"editor.js",
		"layout.js",
		"render.js",
		"ops.js",
		"portions.js"
	],
	remainingModules: [
		"autocorrect.js",
		"drawing.js",
		"drawtool.js",
		"fields.js",
		"mailmerge.js",
		"omml.js",
		"ops.js",
		"outline.js",
		"panes.js",
		"portions.js",
		"preserve-drawing.js",
		"preserve.js",
		"review.js",
		"rtf.js",
		"spell.js",
		"tables.js",
		"templates.js"
	]
}, D = [
	["Typing and selecting", "Click to place the insertion point and type. Double-click selects a word; Ctrl+A selects the document."],
	["Formatting text", "Use the Home ribbon to format text. Ctrl+B, Ctrl+I and Ctrl+U are available in the editor."],
	["Files", "Open and save Office Open XML, Markdown, HTML and plain text documents."]
];
function O({ text: e, onNew: t, onOpen: n, onClose: r }) {
	let [i, a] = h("home"), [o, s] = h(""), c = p(() => D.filter(([e, t]) => `${e} ${t}`.toLowerCase().includes(o.toLowerCase())), [o]);
	return /* @__PURE__ */ x("aside", {
		className: "vo-task-pane",
		"aria-label": "Quire task pane",
		children: [
			/* @__PURE__ */ x("header", { children: [
				/* @__PURE__ */ b("button", {
					type: "button",
					onClick: () => a("home"),
					children: "Home"
				}),
				/* @__PURE__ */ b("strong", { children: i === "help" ? "Help" : i === "search" ? "Search" : "Getting Started" }),
				/* @__PURE__ */ b("button", {
					type: "button",
					onClick: r,
					"aria-label": "Close task pane",
					children: "×"
				})
			] }),
			i === "home" && /* @__PURE__ */ x("div", {
				className: "vo-task-body",
				children: [
					/* @__PURE__ */ b("button", {
						type: "button",
						onClick: n,
						children: "Open a document"
					}),
					/* @__PURE__ */ b("button", {
						type: "button",
						onClick: t,
						children: "Create a blank document"
					}),
					/* @__PURE__ */ b("button", {
						type: "button",
						onClick: () => a("help"),
						children: "Keyboard shortcuts and help"
					}),
					/* @__PURE__ */ b("button", {
						type: "button",
						onClick: () => a("search"),
						children: "Search this document"
					})
				]
			}),
			i === "help" && /* @__PURE__ */ b("div", {
				className: "vo-task-body",
				children: D.map(([e, t]) => /* @__PURE__ */ x("section", { children: [/* @__PURE__ */ b("h4", { children: e }), /* @__PURE__ */ b("p", { children: t })] }, e))
			}),
			i === "search" && /* @__PURE__ */ x("div", {
				className: "vo-task-body",
				children: [
					/* @__PURE__ */ b("input", {
						autoFocus: !0,
						value: o,
						onChange: (e) => s(e.target.value),
						placeholder: "Search for"
					}),
					o && /* @__PURE__ */ b("p", { children: e.toLowerCase().includes(o.toLowerCase()) ? "Match found in this document." : "No matches were found in this document." }),
					c.map(([e, t]) => /* @__PURE__ */ x("section", { children: [/* @__PURE__ */ b("h4", { children: e }), /* @__PURE__ */ b("p", { children: t })] }, e))
				]
			})
		]
	});
}
// src/components/EditorSurface.tsx
function k({ value: e, readOnly: t, onChange: n, onContextMenu: r }) {
	let i = m(null);
	return /* @__PURE__ */ b("main", {
		className: "vo-editor",
		children: /* @__PURE__ */ b("textarea", {
			ref: i,
			readOnly: t,
			value: e,
			onChange: (e) => n(e.target.value),
			onContextMenu: (e) => {
				e.preventDefault(), r?.(e.clientX, e.clientY);
			},
			"aria-label": "Document editor"
		})
	});
}
// src/components/Ruler.tsx
function A({ orientation: e = "horizontal", zoom: t = 1, length: n = 720 }) {
	let r = m(null);
	return f(() => {
		let i = r.current;
		if (!i) return;
		let a = window.devicePixelRatio || 1, o = e === "horizontal" ? n : 28, s = e === "horizontal" ? 24 : n;
		i.width = o * a, i.height = s * a, i.style.width = `${o}px`, i.style.height = `${s}px`;
		let c = i.getContext("2d");
		if (!c) return;
		c.setTransform(a, 0, 0, a, 0, 0), c.clearRect(0, 0, o, s), c.fillStyle = "#f5f5f5", c.fillRect(0, 0, o, s), c.strokeStyle = "#777", c.fillStyle = "#333", c.font = "9px sans-serif";
		let l = 72 * t;
		for (let t = 0, r = 0; t < n; t += l / 8, r++) {
			let n = r % 8 == 0;
			c.beginPath(), e === "horizontal" ? (c.moveTo(t, s), c.lineTo(t, s - (n ? 10 : 4)), n && r && c.fillText(String(r / 8), t + 2, 10)) : (c.moveTo(o, t), c.lineTo(o - (n ? 10 : 4), t), n && r && c.fillText(String(r / 8), 2, t + 10)), c.stroke();
		}
	}, [
		e,
		t,
		n
	]), /* @__PURE__ */ b("canvas", {
		ref: r,
		className: `vo-ruler vo-ruler-${e}`,
		"aria-label": `${e} ruler`
	});
}
// src/components/MailMergePanel.tsx
function j({ onClose: e, onSelect: t }) {
	let [n, r] = h(null), [i, o] = h(""), s = p(() => n?.rows.filter((e) => !i || e.join(" ").toLowerCase().includes(i.toLowerCase())) ?? [], [n, i]);
	return /* @__PURE__ */ x("aside", {
		className: "vo-task-pane",
		"aria-label": "Mail Merge",
		children: [/* @__PURE__ */ x("header", { children: [/* @__PURE__ */ b("strong", { children: "Mail Merge" }), /* @__PURE__ */ b("button", {
			type: "button",
			onClick: e,
			children: "×"
		})] }), /* @__PURE__ */ x("div", {
			className: "vo-task-body",
			children: [
				/* @__PURE__ */ b("input", {
					type: "file",
					accept: ".csv,.tsv,.txt",
					onChange: async (e) => {
						let t = e.target.files?.[0];
						t && r(a(await t.text()));
					}
				}),
				/* @__PURE__ */ b("input", {
					placeholder: "Filter recipients",
					value: i,
					onChange: (e) => o(e.target.value)
				}),
				n && /* @__PURE__ */ x("table", { children: [/* @__PURE__ */ b("thead", { children: /* @__PURE__ */ b("tr", { children: n.fields.map((e) => /* @__PURE__ */ b("th", { children: e }, e)) }) }), /* @__PURE__ */ b("tbody", { children: s.map((e, r) => /* @__PURE__ */ b("tr", {
					onClick: () => t?.(Object.fromEntries(n.fields.map((t, n) => [t, e[n] ?? ""]))),
					children: e.map((e, t) => /* @__PURE__ */ b("td", { children: e }, t))
				}, r)) })] })
			]
		})]
	});
}
// src/components/ReviewPanel.tsx
function M({ revisions: e = [], comments: t = [], onRevision: n, onClose: r }) {
	let [i, a] = h("changes");
	return /* @__PURE__ */ x("aside", {
		className: "vo-task-pane",
		"aria-label": "Review",
		children: [/* @__PURE__ */ x("header", { children: [/* @__PURE__ */ b("strong", { children: "Review" }), /* @__PURE__ */ b("button", {
			type: "button",
			onClick: r,
			children: "×"
		})] }), /* @__PURE__ */ x("div", {
			className: "vo-task-body",
			children: [/* @__PURE__ */ x("nav", { children: [/* @__PURE__ */ x("button", {
				type: "button",
				onClick: () => a("changes"),
				children: [
					"Changes (",
					e.length,
					")"
				]
			}), /* @__PURE__ */ x("button", {
				type: "button",
				onClick: () => a("comments"),
				children: [
					"Comments (",
					t.length,
					")"
				]
			})] }), i === "changes" ? e.map((e) => /* @__PURE__ */ x("section", { children: [
				/* @__PURE__ */ x("p", { children: [
					e.kind,
					": ",
					e.text
				] }),
				/* @__PURE__ */ b("button", {
					type: "button",
					onClick: () => n?.({
						...e,
						accepted: !0
					}),
					children: "Accept"
				}),
				/* @__PURE__ */ b("button", {
					type: "button",
					onClick: () => n?.({
						...e,
						accepted: !1
					}),
					children: "Reject"
				})
			] }, e.id)) : t.map((e) => /* @__PURE__ */ x("section", { children: [/* @__PURE__ */ b("strong", { children: e.author || "Comment" }), /* @__PURE__ */ b("p", { children: e.text })] }, e.id))]
		})]
	});
}
// src/components/RWordWorkspace.tsx
function N({ initialPanel: e = "task", ...t }) {
	let n = m(null), [r, i] = h(e);
	return /* @__PURE__ */ x("div", {
		className: "vo-word-workspace",
		children: [
			/* @__PURE__ */ x("div", {
				className: "vo-word-main",
				children: [/* @__PURE__ */ b(R, {
					ref: n,
					...t
				}), /* @__PURE__ */ x("div", {
					className: "vo-workspace-actions",
					children: [
						/* @__PURE__ */ b("button", {
							type: "button",
							onClick: () => i(r === "task" ? "none" : "task"),
							children: "Task Pane"
						}),
						/* @__PURE__ */ b("button", {
							type: "button",
							onClick: () => i(r === "mailmerge" ? "none" : "mailmerge"),
							children: "Mail Merge"
						}),
						/* @__PURE__ */ b("button", {
							type: "button",
							onClick: () => i(r === "review" ? "none" : "review"),
							children: "Review"
						})
					]
				})]
			}),
			r === "task" && /* @__PURE__ */ b(O, {
				text: "",
				onNew: () => n.current?.newDocument(),
				onOpen: () => void 0,
				onClose: () => i("none")
			}),
			r === "mailmerge" && /* @__PURE__ */ b(j, { onClose: () => i("none") }),
			r === "review" && /* @__PURE__ */ b(M, { onClose: () => i("none") })
		]
	});
}
// src/components/OutlinePanel.tsx
function P({ doc: e, onSelect: t, onClose: n }) {
	let r = p(() => c(e), [e]), i = (e) => /* @__PURE__ */ b("ul", { children: e.map((e) => /* @__PURE__ */ x("li", { children: [/* @__PURE__ */ b("button", {
		type: "button",
		onClick: () => t?.(e),
		children: e.title
	}), e.children.length ? i(e.children) : null] }, e.id)) });
	return /* @__PURE__ */ x("aside", {
		className: "vo-task-pane",
		"aria-label": "Document Map",
		children: [/* @__PURE__ */ x("header", { children: [/* @__PURE__ */ b("strong", { children: "Document Map" }), /* @__PURE__ */ b("button", {
			type: "button",
			onClick: n,
			children: "×"
		})] }), /* @__PURE__ */ b("div", {
			className: "vo-task-body",
			children: r.length ? i(r) : /* @__PURE__ */ b("p", { children: "(No headings)" })
		})]
	});
}
// src/components/FormattingPanel.tsx
function F({ onApply: e, onClose: t }) {
	return /* @__PURE__ */ x("aside", {
		className: "vo-task-pane",
		"aria-label": "Formatting",
		children: [/* @__PURE__ */ x("header", { children: [/* @__PURE__ */ b("strong", { children: "Formatting" }), /* @__PURE__ */ b("button", {
			type: "button",
			onClick: t,
			children: "×"
		})] }), /* @__PURE__ */ x("div", {
			className: "vo-task-body",
			children: [
				/* @__PURE__ */ b("button", {
					type: "button",
					onClick: () => e?.({ bold: !0 }),
					children: "Bold"
				}),
				/* @__PURE__ */ b("button", {
					type: "button",
					onClick: () => e?.({ italic: !0 }),
					children: "Italic"
				}),
				/* @__PURE__ */ b("button", {
					type: "button",
					onClick: () => e?.({ underline: !0 }),
					children: "Underline"
				}),
				/* @__PURE__ */ b("button", {
					type: "button",
					onClick: () => e?.({}),
					children: "Clear formatting"
				})
			]
		})]
	});
}
// src/components/CommandSurface.tsx
function I({ menus: e, tools: t }) {
	return /* @__PURE__ */ x(y, { children: [/* @__PURE__ */ b(ie, { menus: e }), /* @__PURE__ */ b(_, { items: t })] });
}
// src/index.tsx
function L(e, t = "docx") {
	let n = l();
	return n.format = t, n.blocks = e.split(/\r?\n/).map((e) => ({
		...s(),
		runs: e ? [{ text: e }] : []
	})), n;
}
var R = d(function({ src: a, fileName: s = "Document.docx", readonly: c, onReady: d, onChange: p, onSave: y, onError: S }, C) {
	let [w, T] = h(s), [E, D] = h(!1), [O, k] = h(""), [j, M] = h(""), [N, P] = h(null), [F, I] = h(null), [R, fe] = h(""), [z, pe] = h(""), [B, me] = h(3), [he, ge] = h(3), [_e, ve] = h(!0), V = m(l()), H = m(null), ye = m("word/document.xml"), U = m(null), W = m([]), G = m([]), K = (e) => {
		D(e), p?.(e);
	};
	ue(E);
	let q = async (e, t) => {
		try {
			k("Opening…");
			let r = e instanceof Blob ? {
				blob: e,
				name: t ?? (e instanceof File ? e.name : "Document.docx")
			} : await le(typeof e == "string" ? {
				kind: "http",
				url: e
			} : e), a = r.name.split(".").pop()?.toLowerCase(), o = new Uint8Array(await r.blob.arrayBuffer());
			if (a === "md" || a === "markdown") V.current = L(n(new TextDecoder().decode(o)).text, "md"), H.current = null, M(u(V.current));
			else if (a === "html" || a === "htm") V.current = L(i(new TextDecoder().decode(o)), "html"), H.current = null, M(u(V.current));
			else {
				let e = await ee(o);
				V.current = e.doc, H.current = e.pkg, ye.current = e.part, M(u(e.doc));
			}
			U.current = e, T(r.name), K(!1), d?.();
		} catch (e) {
			S?.(e instanceof Error ? e : Error(String(e)));
		} finally {
			k("");
		}
	}, J = async () => V.current.format === "md" ? new Blob([r(j)], { type: "text/markdown;charset=utf-8" }) : V.current.format === "html" ? new Blob([o(j, w)], { type: "text/html;charset=utf-8" }) : e(V.current, {
		pkg: H.current ?? void 0,
		part: ye.current,
		format: V.current.format
	}), be = async (e = w) => {
		let t = await J(), n = /\.[a-z0-9]+$/i.test(e) ? e : `${e}.docx`;
		return ce(t, n), K(!1), t;
	}, Y = async () => {
		let e = await J(), t = U.current;
		t && typeof t != "string" && "kind" in t && t.kind !== "http" ? await v(t, e, w) : y?.(e, w), K(!1);
	}, X = () => {
		V.current = l(), H.current = null, U.current = null, W.current = [], G.current = [], M(""), T("Document.docx"), K(!1);
	};
	te(C, () => ({
		open: q,
		newDocument: X,
		getBlob: J,
		save: Y,
		saveAs: be,
		saveTo: async (e, t) => v(e, await J(), t ?? w),
		get name() {
			return w;
		},
		get dirty() {
			return E;
		}
	}), [
		w,
		E,
		j
	]), f(() => {
		a && q(a);
	}, []), f(() => {
		let e = (e) => {
			if (!(e.ctrlKey || e.metaKey)) return;
			let t = e.key.toLowerCase();
			t === "s" ? (e.preventDefault(), Y()) : t === "z" ? (e.preventDefault(), e.shiftKey ? $() : Q()) : t === "y" ? (e.preventDefault(), $()) : t === "n" ? (e.preventDefault(), X()) : t === "o" ? (e.preventDefault(), document.querySelector("#rword-file")?.click()) : t === "f" && (e.preventDefault(), I("find"));
		};
		return window.addEventListener("keydown", e), () => window.removeEventListener("keydown", e);
	}, [
		j,
		w,
		E
	]);
	let Z = (e) => {
		e !== j && (W.current.push(j), G.current = [], W.current.length > 50 && W.current.shift()), M(e), V.current = L(e, V.current.format), K(!0);
	}, xe = () => {
		if (!R) return;
		let e = de(j, {
			text: R,
			replace: z
		});
		e.count && Z(e.value);
	}, Se = () => {
		t(V.current, B, he), M(u(V.current)), K(!0), I(null);
	}, Q = () => {
		let e = W.current.pop();
		e !== void 0 && (G.current.push(j), M(e), V.current = L(e, V.current.format), K(!0));
	}, $ = () => {
		let e = G.current.pop();
		e !== void 0 && (W.current.push(j), M(e), V.current = L(e, V.current.format), K(!0));
	}, Ce = [
		{
			label: "File",
			items: [
				{
					label: "New",
					key: "Ctrl+N",
					run: X
				},
				{
					label: "Open…",
					key: "Ctrl+O",
					run: () => document.querySelector("#rword-file")?.click()
				},
				{
					label: "Open from Online…",
					run: () => I("open")
				},
				{
					sep: !0,
					label: ""
				},
				{
					label: "Save",
					key: "Ctrl+S",
					run: Y
				},
				{
					label: "Save As…",
					run: async () => {
						await be();
					}
				},
				{
					label: "Save to Online…",
					run: () => I("save")
				},
				{
					label: "Print…",
					key: "Ctrl+P",
					run: () => window.print()
				}
			]
		},
		{
			label: "Edit",
			items: [
				{
					label: "Undo",
					key: "Ctrl+Z",
					run: Q
				},
				{
					label: "Redo",
					key: "Ctrl+Y",
					run: $
				},
				{
					sep: !0,
					label: ""
				},
				{
					label: "Find…",
					key: "Ctrl+F",
					run: () => I("find")
				},
				{
					label: "Replace…",
					run: () => I("find")
				},
				{
					label: "Select All",
					key: "Ctrl+A",
					run: () => document.querySelector(".vo-editor textarea")?.select()
				}
			]
		},
		{
			label: "View",
			items: [
				{
					label: "Print Layout",
					checked: !0,
					run: () => void 0
				},
				{
					label: "Zoom In",
					run: () => void 0
				},
				{
					label: "Zoom Out",
					run: () => void 0
				}
			]
		},
		{
			label: "Insert",
			items: [
				{
					label: "Page Break",
					run: () => Z(j + "\n\n")
				},
				{
					label: "Table…",
					run: () => I("table")
				},
				{
					label: "Date and Time",
					run: () => Z(j + (/* @__PURE__ */ new Date()).toLocaleDateString())
				},
				{
					label: "Symbol…",
					disabled: !0
				}
			]
		},
		{
			label: "Format",
			items: [
				{
					label: "Bold",
					key: "Ctrl+B",
					disabled: !0
				},
				{
					label: "Italic",
					key: "Ctrl+I",
					disabled: !0
				},
				{
					label: "Underline",
					key: "Ctrl+U",
					disabled: !0
				}
			]
		},
		{
			label: "Tools",
			items: [{
				label: "Word Count",
				run: () => void 0
			}, {
				label: "About Quire",
				run: () => void 0
			}]
		},
		{
			label: "Help",
			items: [{
				label: "About Quire",
				run: () => I("about")
			}]
		}
	], we = [
		{
			id: "new",
			label: "New",
			title: "New",
			run: X
		},
		{
			id: "open",
			label: "Open",
			title: "Open",
			run: () => document.querySelector("#rword-file")?.click()
		},
		{
			id: "save",
			label: "Save",
			title: "Save",
			run: Y
		},
		{
			id: "undo",
			label: "Undo",
			title: "Undo",
			run: Q
		},
		{
			id: "redo",
			label: "Redo",
			title: "Redo",
			run: $
		},
		{
			id: "print",
			label: "Print",
			title: "Print",
			run: () => window.print()
		},
		{
			id: "sep",
			kind: "sep"
		},
		{
			id: "find",
			label: "Find",
			title: "Find",
			run: () => I("find")
		}
	];
	return /* @__PURE__ */ b(re, { children: /* @__PURE__ */ x("div", {
		className: "vo-office",
		children: [
			/* @__PURE__ */ x("div", {
				className: "vo-title",
				children: [
					/* @__PURE__ */ b("span", {
						className: "vo-title-app",
						children: "Quire"
					}),
					" — ",
					w,
					E ? " *" : ""
				]
			}),
			/* @__PURE__ */ b("input", {
				id: "rword-file",
				hidden: !0,
				type: "file",
				accept: ".docx,.docm,.dotx,.dotm,.md,.markdown,.html,.htm,.txt",
				onChange: (e) => {
					let t = e.target.files?.[0];
					t && q(t);
				}
			}),
			/* @__PURE__ */ b(ie, { menus: Ce }),
			/* @__PURE__ */ x("div", {
				id: "toolbars",
				className: "vo-toolbars",
				children: [/* @__PURE__ */ b(_, { items: we }), /* @__PURE__ */ b(_, { items: [
					{
						id: "bold",
						label: "Bold",
						title: "Bold",
						run: () => {
							document.execCommand("bold");
						}
					},
					{
						id: "italic",
						label: "Italic",
						title: "Italic",
						run: () => {
							document.execCommand("italic");
						}
					},
					{
						id: "underline",
						label: "Underline",
						title: "Underline",
						run: () => {
							document.execCommand("underline");
						}
					},
					{
						id: "sep-format",
						kind: "sep"
					},
					{
						id: "left",
						label: "Align Left",
						title: "Align Left",
						run: () => {
							document.execCommand("justifyLeft");
						}
					},
					{
						id: "center",
						label: "Center",
						title: "Center",
						run: () => {
							document.execCommand("justifyCenter");
						}
					},
					{
						id: "right",
						label: "Align Right",
						title: "Align Right",
						run: () => {
							document.execCommand("justifyRight");
						}
					}
				] })]
			}),
			/* @__PURE__ */ b(oe, { tabs: [
				{
					id: "home",
					label: "Home",
					items: we.slice(0, 4)
				},
				{
					id: "insert",
					label: "Insert",
					items: [
						{
							id: "break",
							label: "Page Break",
							run: () => Z(j + "\n\n")
						},
						{
							id: "table",
							label: "Table",
							run: () => I("table")
						},
						{
							id: "date",
							label: "Date and Time",
							run: () => Z(j + (/* @__PURE__ */ new Date()).toLocaleDateString())
						}
					]
				},
				{
					id: "view",
					label: "View",
					items: [{
						id: "print",
						label: "Print Layout"
					}, {
						id: "task",
						label: "Task Pane",
						run: () => ve((e) => !e)
					}]
				}
			] }),
			/* @__PURE__ */ x("div", {
				className: "vo-main",
				children: [/* @__PURE__ */ b("div", {
					className: "vo-center",
					children: /* @__PURE__ */ x("div", {
						className: "vo-work",
						children: [
							/* @__PURE__ */ b("div", { className: "vo-ruler-corner" }),
							/* @__PURE__ */ b("div", {
								className: "vo-ruler-h",
								children: /* @__PURE__ */ b(A, {
									orientation: "horizontal",
									length: 900
								})
							}),
							/* @__PURE__ */ b("div", {
								className: "vo-ruler-v",
								children: /* @__PURE__ */ b(A, {
									orientation: "vertical",
									length: 900
								})
							}),
							/* @__PURE__ */ b("main", {
								className: "vo-editor",
								children: /* @__PURE__ */ b("textarea", {
									readOnly: c,
									value: j,
									onContextMenu: (e) => {
										e.preventDefault(), P({
											x: e.clientX,
											y: e.clientY
										});
									},
									onChange: (e) => Z(e.target.value),
									"aria-label": "Document editor"
								})
							})
						]
					})
				}), _e && /* @__PURE__ */ x("aside", {
					className: "vo-taskpane",
					children: [/* @__PURE__ */ b("div", {
						className: "vo-taskpane-title",
						children: "Document"
					}), /* @__PURE__ */ x("div", {
						className: "vo-taskpane-body",
						children: [
							/* @__PURE__ */ b("button", {
								type: "button",
								onClick: X,
								children: "New document"
							}),
							/* @__PURE__ */ b("button", {
								type: "button",
								onClick: () => document.querySelector("#rword-file")?.click(),
								children: "Open document"
							}),
							/* @__PURE__ */ b("p", { children: j ? `${j.trim().split(/\s+/).filter(Boolean).length} words` : "Start typing to begin." })
						]
					})]
				})]
			}),
			/* @__PURE__ */ x(se, { children: [
				/* @__PURE__ */ x("span", { children: [j.trim().split(/\s+/).filter(Boolean).length, " words"] }),
				/* @__PURE__ */ b("span", {
					className: "vo-status-message",
					children: O || "Ready"
				}),
				/* @__PURE__ */ b("button", {
					type: "button",
					onClick: () => ve((e) => !e),
					children: "Task Pane"
				})
			] }),
			N && /* @__PURE__ */ b(ne, {
				x: N.x,
				y: N.y,
				onClose: () => P(null),
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
						label: "Select All",
						run: () => document.querySelector(".vo-editor textarea")?.select()
					}
				]
			}),
			F === "open" && /* @__PURE__ */ b(ae, {
				mode: "open",
				onClose: () => I(null),
				onSubmit: (e) => {
					I(null), q(e);
				}
			}),
			F === "save" && /* @__PURE__ */ b(ae, {
				mode: "save",
				fileName: w,
				onClose: () => I(null),
				onSubmit: (e) => {
					I(null), v({
						kind: "presigned",
						url: e,
						putUrl: e
					}, J(), w);
				}
			}),
			F === "about" && /* @__PURE__ */ b(g, {
				title: "About Quire",
				onClose: () => I(null),
				children: /* @__PURE__ */ b("p", { children: "Quire is the VibeOffice word processor." })
			}),
			F === "find" && /* @__PURE__ */ x(g, {
				title: "Find and Replace",
				onClose: () => I(null),
				children: [
					/* @__PURE__ */ x("label", { children: ["Find", /* @__PURE__ */ b("input", {
						value: R,
						onChange: (e) => fe(e.target.value)
					})] }),
					/* @__PURE__ */ x("label", { children: ["Replace with", /* @__PURE__ */ b("input", {
						value: z,
						onChange: (e) => pe(e.target.value)
					})] }),
					/* @__PURE__ */ x("div", { children: [/* @__PURE__ */ b("button", {
						onClick: () => {
							let e = j.indexOf(R);
							e >= 0 && document.querySelector(".vo-editor textarea")?.setSelectionRange(e, e + R.length);
						},
						children: "Find Next"
					}), /* @__PURE__ */ b("button", {
						onClick: xe,
						children: "Replace"
					})] })
				]
			}),
			F === "table" && /* @__PURE__ */ x(g, {
				title: "Insert Table",
				onClose: () => I(null),
				children: [
					/* @__PURE__ */ x("label", { children: ["Rows", /* @__PURE__ */ b("input", {
						type: "number",
						min: "1",
						max: "20",
						value: B,
						onChange: (e) => me(Math.max(1, Math.min(20, Number(e.target.value))))
					})] }),
					/* @__PURE__ */ x("label", { children: ["Columns", /* @__PURE__ */ b("input", {
						type: "number",
						min: "1",
						max: "20",
						value: he,
						onChange: (e) => ge(Math.max(1, Math.min(20, Number(e.target.value))))
					})] }),
					/* @__PURE__ */ b("button", {
						onClick: Se,
						children: "Insert"
					})
				]
			})
		]
	}) });
});
export { I as CommandSurface, k as EditorSurface, F as FormattingPanel, j as MailMergePanel, P as OutlinePanel, R as RWord, R as default, N as RWordWorkspace, M as ReviewPanel, A as Ruler, O as TaskPane, E as quireScriptMigration };
