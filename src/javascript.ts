import { Length } from "./document/common";
import { ParagraphTab } from "./document/paragraph";

interface TabStop {
	pos: number;
	leader: string;
	style: string;
}

const defaultTab: TabStop = { pos: 0, leader: "none", style: "left" };
const maxTabs = 50;

export function computePixelToPoint(container: HTMLElement = document.body) {
	const temp = document.createElement("div");
	temp.style.width = '100pt';
	
	container.appendChild(temp);
	const result = 100 / temp.offsetWidth;
	container.removeChild(temp);

	return result
}

export function updateTabStop(elem: HTMLElement, tabs: ParagraphTab[], defaultTabSize: Length, pixelToPoint: number = 72 / 96) {
    const p = elem.closest("p");
    if (!p || !p.offsetWidth) return;
    // DOM rectangles include CSS zoom/transforms, while tab positions are points.
    const pbb = p.getBoundingClientRect();
    const scale = pbb.width / p.offsetWidth || 1;
    const ratio = pixelToPoint / scale;
    const left = (elem.getBoundingClientRect().left - pbb.left) * ratio;
    const stops: TabStop[] = (tabs || []).filter(t => t.style !== "clear").map(t => ({
        pos: lengthToPoint(t.position), leader: t.leader, style: t.style
    })).sort((a,b) => a.pos - b.pos);
    const size = lengthToPoint(defaultTabSize) || 36;
    let pos = (stops[stops.length - 1]?.pos || 0) + size;
    for (; pos <= p.offsetWidth * pixelToPoint && stops.length < maxTabs; pos += size) {
        stops.push({...defaultTab, pos});
    }
    const tab = stops.find(t => t.pos > left + 0.1);
    if (!tab) return;
    let following = 0;
    if (tab.style === "right" || tab.style === "center") {
        const range = document.createRange();
        range.setStartAfter(elem);
        const next = Array.from(p.querySelectorAll("[data-docx-tab]")).find(t =>
            !!(elem.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_FOLLOWING));
        if (next) range.setEndBefore(next); else range.setEnd(p, p.childNodes.length);
        // Empty ranges have a zero rect at the viewport origin. They occupy no
        // following text and must not be measured relative to that origin.
        if (range.toString().trim()) following = range.getBoundingClientRect().width * ratio;
    }
    const width = Math.max(0, tab.pos - left - following * (tab.style === "center" ? .5 : 1));
    delete elem.dataset.tabLayout;
    elem.textContent = "x";
    Object.assign(elem.style, {display:"inline-block",width:`${width}pt`,wordSpacing:"0",letterSpacing:"0",whiteSpace:"pre",tabSize:"0",fontSize:"inherit",minHeight:"0",textDecoration:"inherit"});
    const glyph = document.createRange();
    glyph.selectNodeContents(elem);
    const glyphWidth = glyph.getBoundingClientRect().width * ratio;
    // Keep a real tab in selectable text. Its CSS advance is zero; the native
    // decoration carrier in ::after supplies the exact measured advance.
    elem.textContent = "\t";
    elem.style.setProperty('--docx-tab-spacing', `${width - glyphWidth}pt`);
    elem.dataset.tabLayout = 'native';
    // Generated decoration content is visual only; do not announce its carrier.
    elem.setAttribute('aria-hidden', 'true');
    elem.dataset.docxTab = "true";
    if (["dot","middleDot","hyphen","heavy","underscore"].includes(tab.leader)) {
        elem.style.textDecoration = "underline";
        if (["dot","middleDot"].includes(tab.leader)) elem.style.textDecorationStyle = "dotted";
    }
}

function lengthToPoint(length: Length): number {
	return parseFloat(length);
}
