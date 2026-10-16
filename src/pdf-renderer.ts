/** An authoritative DOCX engine supplies PDF; this renderer preserves its pages. */
export interface PdfPage {
    width: number;
    height: number;
    imageUrl: string;
}

export interface PdfPreviewResult {
    pdfBytes: Uint8Array;
    pageCount: number;
}

// Keep the PDF engine outside the core bundle and let applications choose their
// deployment/worker arrangement. No uploads or third-party services are implied.
export interface PdfPreviewAdapter {
    convert(data: Blob | ArrayBuffer | Uint8Array): Promise<Uint8Array>;
    renderPages(pdf: Uint8Array): Promise<PdfPage[]>;
}

export async function renderPdfPreview(data: Blob | ArrayBuffer | Uint8Array, body: HTMLElement, styles: HTMLElement, adapter: PdfPreviewAdapter, className: string): Promise<PdfPreviewResult> {
    if (!adapter) throw new Error('PDF preview requires an authoritative PDF adapter');
    const pdfBytes = await adapter.convert(data);
    const pages = await adapter.renderPages(pdfBytes.slice());
    if (!pages.length) throw new Error('PDF preview contains no pages');
    const wrapper = document.createElement('div');
    wrapper.className = `${className}-wrapper`;
    wrapper.style.cssText = 'display:flex;flex-direction:column;align-items:center;gap:14px';
    for (const [index,page] of pages.entries()) {
        if (!(Number.isFinite(page.width) && Number.isFinite(page.height) && page.width > 0 && page.height > 0) || !page.imageUrl.startsWith('data:image/')) throw new Error('Invalid PDF page');
        const section = document.createElement('section');
        section.className = className;
        section.dataset.pageNumber = String(index + 1);
        section.style.cssText = `width:${page.width}px;height:${page.height}px;flex:none;background:white;box-shadow:0 2px 8px #0003`;
        const image = document.createElement('img');
        image.src = page.imageUrl;
        image.alt = `Page ${index + 1} of ${pages.length}`;
        image.style.cssText = 'display:block;width:100%;height:100%';
        section.append(image);
        wrapper.append(section);
    }
    await Promise.all(Array.from(wrapper.querySelectorAll('img')).map(image=>image.decode()));
    styles.replaceChildren();
    body.replaceChildren(wrapper);
    body.dataset.renderMode = 'pdf';
    return {pdfBytes, pageCount:pages.length};
}
