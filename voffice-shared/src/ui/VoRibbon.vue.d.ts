import type { RibbonTab } from './types';
type __VLS_Props = {
    tabs: RibbonTab[];
    modelValue?: string;
};
declare const __VLS_export: import("vue").DefineComponent<__VLS_Props, {}, {}, {}, {}, import("vue").ComponentOptionsMixin, import("vue").ComponentOptionsMixin, {
    file: () => any;
    "update:modelValue": (args_0: string) => any;
}, string, import("vue").PublicProps, Readonly<__VLS_Props> & Readonly<{
    onFile?: (() => any) | undefined;
    "onUpdate:modelValue"?: ((args_0: string) => any) | undefined;
}>, {}, {}, {}, {}, string, import("vue").ComponentProvideOptions, false, {}, any>;
declare const _default: typeof __VLS_export;
export default _default;
