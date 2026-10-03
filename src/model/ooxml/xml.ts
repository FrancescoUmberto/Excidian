import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

export const NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export function parseXml(text: string): Document {
	return new DOMParser().parseFromString(text, "text/xml") as unknown as Document;
}

export function serializeXml(doc: Document): string {
	const xml = new XMLSerializer().serializeToString(doc as never);
	return xml.startsWith("<?xml") ? xml : XML_DECLARATION + xml;
}

/** Child elements of `parent`, optionally only those with the given local name. */
export function elements(parent: Node, localName?: string): Element[] {
	const out: Element[] = [];
	for (let n = parent.firstChild; n; n = n.nextSibling) {
		if (n.nodeType === 1 && (!localName || (n as Element).localName === localName)) out.push(n as Element);
	}
	return out;
}

export function child(parent: Node, localName: string): Element | undefined {
	for (let n = parent.firstChild; n; n = n.nextSibling) {
		if (n.nodeType === 1 && (n as Element).localName === localName) return n as Element;
	}
	return undefined;
}

/** Creates an element in the parent's namespace (and prefix) and inserts it before `before`. */
export function createChild(parent: Element, localName: string, before: Node | null = null): Element {
	const name = parent.prefix ? `${parent.prefix}:${localName}` : localName;
	const el = parent.ownerDocument.createElementNS(parent.namespaceURI, name);
	parent.insertBefore(el, before);
	return el;
}

export function textOf(el: Element | undefined): string {
	return el?.textContent ?? "";
}

export function setText(el: Element, text: string) {
	while (el.firstChild) el.removeChild(el.firstChild);
	el.appendChild(el.ownerDocument.createTextNode(text));
}

export function removeChildren(el: Element) {
	while (el.firstChild) el.removeChild(el.firstChild);
}

/** Text of a shared string or inline string: plain `<t>` or rich-text runs, without phonetic hints. */
export function richText(el: Element): string {
	let text = "";
	for (const part of elements(el)) {
		if (part.localName === "t") text += textOf(part);
		else if (part.localName === "r") text += textOf(child(part, "t"));
	}
	return text;
}
