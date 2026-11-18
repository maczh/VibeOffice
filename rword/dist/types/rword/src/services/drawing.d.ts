import type { ImageRef, Run } from '../engine/model';
export declare function imageRun(image: ImageRef, alt?: string): Run;
export declare function imageSize(width: number, height: number, maxWidth?: number): {
    width: number;
    height: number;
};
export declare function mediaType(name: string): string;
