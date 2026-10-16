import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';
import { h } from './html';
import { renderPdfPreview, PdfPreviewAdapter } from './pdf-renderer';
export type { PdfPage, PdfPreviewAdapter, PdfPreviewResult } from './pdf-renderer';

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
    h: typeof h;
}

export const defaultOptions: Options = {
    ignoreHeight: false,
    ignoreWidth: false,
    ignoreFonts: false,
    breakPages: true,
    debug: false,
    experimental: false,
    className: "docx",
    inWrapper: true,
    hideWrapperOnPrint: false,
    trimXmlDeclaration: true,
    ignoreLastRenderedPageBreak: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
	renderEndnotes: true,
	useBase64URL: false,
	renderChanges: false,
    renderComments: false,
    renderAltChunks: true,
    h: h
};

export function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<any>  {
    const ops = { ...defaultOptions, ...userOptions };
    return WordDocument.load(data, new DocumentParser(ops), ops);
}

export async function renderDocument(document: any, userOptions?: Partial<Options>): Promise<any> {
    const ops = { ...defaultOptions, ...userOptions };
    const renderer = new HtmlRenderer();
    return await renderer.render(document, ops);
}

export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any> {
	if (userOptions?.renderMode === 'pdf') {
        return renderPdfPreview(data, bodyContainer, styleContainer ?? document.createElement('div'), userOptions.pdfAdapter, userOptions.className ?? 'docx');
    }
	const doc = await parseAsync(data, userOptions);
	const renderer = new HtmlRenderer();
	const nodes = await renderer.render(doc, {...defaultOptions, ...userOptions});

    styleContainer ??= bodyContainer;
    styleContainer.innerHTML = "";
    bodyContainer.innerHTML = "";

    for (let n of nodes) {
        const c = n.nodeName === "STYLE" ? styleContainer : bodyContainer;
        c.appendChild(n);
    }
    // Detached callers still need connected, unscaled geometry for tab stops.
    let stage: HTMLElement;
    let bodyPlace: Comment, stylePlace: Comment;
    if (userOptions?.experimental && !bodyContainer.isConnected) {
        if (bodyContainer.parentNode) {
            bodyPlace = document.createComment('docx-body');
            bodyContainer.before(bodyPlace);
        }
        if (styleContainer !== bodyContainer && styleContainer.parentNode) {
            stylePlace = document.createComment('docx-styles');
            styleContainer.before(stylePlace);
        }
        stage = document.createElement('div');
        Object.assign(stage.style,{position:'fixed',visibility:'hidden',left:'0',top:'0',width:'max-content'});
        stage.appendChild(bodyContainer);
        if (styleContainer !== bodyContainer) stage.appendChild(styleContainer);
        document.body.appendChild(stage);
    }
    try { await renderer.layout(); }
    finally {
        if (stage) {
            bodyContainer.remove();
            if (styleContainer !== bodyContainer) styleContainer.remove();
            stage.remove();
            bodyPlace?.replaceWith(bodyContainer);
            stylePlace?.replaceWith(styleContainer);
        }
    }
    
    return doc;
}
