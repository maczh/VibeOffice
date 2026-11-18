import { Component, type ReactNode, type ErrorInfo } from 'react';
export * from './office-core';
export * from './chrome';
export * from './common';
export interface WebDavSource {
    kind: 'webdav';
    url: string;
    username?: string;
    password?: string;
    token?: string;
}
export interface S3Source {
    kind: 's3';
    endpoint: string;
    bucket: string;
    key: string;
    region?: string;
    accessKeyId: string;
    secretAccessKey: string;
    sessionToken?: string;
    forcePathStyle?: boolean;
}
export interface PresignedSource {
    kind: 'presigned';
    url: string;
    putUrl?: string;
}
export interface HttpSource {
    kind: 'http';
    url: string;
    headers?: Record<string, string>;
}
export interface HttpPostTarget {
    kind: 'post';
    url: string;
    field?: string;
    fields?: Record<string, string>;
    headers?: Record<string, string>;
}
export type OnlineSource = WebDavSource | S3Source | PresignedSource | HttpSource;
export type SaveTarget = WebDavSource | S3Source | PresignedSource | HttpPostTarget;
export type OpenInput = (File & {
    kind?: undefined;
}) | (Blob & {
    kind?: undefined;
}) | string | OnlineSource;
export interface OnlineFile {
    blob: Blob;
    name: string;
}
export interface OfficeApi {
    open(input: OpenInput, name?: string): Promise<void>;
    newDocument(): void;
    getBlob(): Promise<Blob>;
    save(): Promise<void>;
    saveAs(name?: string): Promise<Blob>;
    saveTo(target: SaveTarget, name?: string): Promise<void>;
    readonly name: string;
    readonly dirty: boolean;
}
export interface OfficeProps {
    src?: OpenInput;
    fileName?: string;
    readonly?: boolean;
}
export declare function nameFromUrl(url: string): string;
export declare function fetchOnline(source: OnlineSource): Promise<OnlineFile>;
export declare function saveOnline(target: SaveTarget, blob: Blob | Promise<Blob>, name: string): Promise<void>;
export declare function downloadBlob(blob: Blob, name: string): void;
export declare class ErrorBoundary extends Component<{
    children: ReactNode;
}, {
    error: Error | null;
}> {
    state: {
        error: Error | null;
    };
    static getDerivedStateFromError(error: Error): {
        error: Error;
    };
    componentDidCatch(error: Error, info: ErrorInfo): void;
    render(): ReactNode;
}
