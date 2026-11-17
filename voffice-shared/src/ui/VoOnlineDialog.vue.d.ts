type __VLS_Props = {
    mode: 'open' | 'save';
    fileName?: string;
};
declare const __VLS_export: import("vue").DefineComponent<__VLS_Props, {}, {}, {}, {}, import("vue").ComponentOptionsMixin, import("vue").ComponentOptionsMixin, {
    close: () => any;
    submit: (source: import("../index.js").WebDavSource | import("../index.js").S3Source | import("../index.js").PresignedSource | import("../index.js").HttpSource | import("../index.js").HttpPostTarget) => any;
}, string, import("vue").PublicProps, Readonly<__VLS_Props> & Readonly<{
    onClose?: (() => any) | undefined;
    onSubmit?: ((source: import("../index.js").WebDavSource | import("../index.js").S3Source | import("../index.js").PresignedSource | import("../index.js").HttpSource | import("../index.js").HttpPostTarget) => any) | undefined;
}>, {
    mode: "open" | "save";
}, {}, {}, {}, string, import("vue").ComponentProvideOptions, false, {}, any>;
declare const _default: typeof __VLS_export;
export default _default;
