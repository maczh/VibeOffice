import { type Ref } from 'vue';
export type Look = 'classic' | 'paper' | 'paper2016';
export declare function getLook(): Look;
export declare function setLook(l: Look): void;
/** a shared reactive look; bind data-look/data-layout on the component root to it */
export declare function useLook(initial?: Look): {
    look: Ref<Look>;
    layout: Ref<string>;
};
