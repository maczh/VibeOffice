import { type Presentation, type Slide, type Shape } from './model';
export declare function shapeHtml(s: Shape): string;
export declare function slideHtml(slide: Slide, pres: Presentation, widthPx: number): string;
