export interface MathRun {
    text: string;
    kind?: 'normal' | 'operator' | 'number';
}
export declare function ommlToMathText(xml: string): string;
export declare function mathTextToOmml(text: string): string;
