export interface RulerProps {
    orientation?: 'horizontal' | 'vertical';
    zoom?: number;
    length?: number;
}
export declare function Ruler({ orientation, zoom, length }: RulerProps): import("react").JSX.Element;
