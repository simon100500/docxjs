/*
 * @license
 * docx-preview <https://github.com/VolodymyrBaydalka/docxjs>
 * Released under Apache License 2.0  <https://github.com/VolodymyrBaydalka/docxjs/blob/master/LICENSE>
 * Copyright Volodymyr Baydalka
 */

export type HElement = {
    ns?: string;
    tagName: string;
    className?: string;
    style?: Record<string, string> | string;
    children?: (HElement | Node | string)[];
} & Record<string, any>;

export interface PdfPage { width:number; height:number; imageUrl:string; }
export interface PdfPreviewResult { pdfBytes:Uint8Array; pageCount:number; }
export interface PdfPreviewAdapter {
 convert(data:Blob|ArrayBuffer|Uint8Array):Promise<Uint8Array>;
 renderPages(pdf:Uint8Array):Promise<PdfPage[]>;
}
export interface Options {
    renderMode?: 'html' | 'pdf';
    pdfAdapter?: PdfPreviewAdapter;
    inWrapper: boolean;
    hideWrapperOnPrint: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
    trimXmlDeclaration: boolean;
    renderHeaders: boolean;
    renderFooters: boolean;
    renderFootnotes: boolean;
    renderEndnotes: boolean;
    ignoreLastRenderedPageBreak: boolean;
    useBase64URL: boolean;
    renderChanges: boolean;
    renderComments: boolean;
    renderAltChunks: boolean;
    h: (elemOrText: HElement | Node | string) => Node; //experimental, subject to change
}

//stub
export type WordDocument = any;
export declare const defaultOptions: Options;
export declare function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<WordDocument>;
export declare function renderDocument(document: WordDocument, userOptions?: Partial<Options>): Promise<Node[]>;
export declare function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any>;
