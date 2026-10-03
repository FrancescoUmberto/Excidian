import JSZip from "jszip";
import { elements, parseXml, serializeXml } from "./xml";

export interface Relationship {
	id: string;
	type: string;
	/** Package path of the target part, without a leading slash. */
	target: string;
}

const CONTENT_TYPES = "[Content_Types].xml";

/**
 * The .xlsx zip container. Parts are parsed on demand, and only parts marked
 * dirty are rewritten on save; everything else is kept byte for byte.
 */
export class Package {
	private docs = new Map<string, Document>();
	private dirty = new Set<string>();

	private constructor(
		private zip: JSZip,
		private texts: Map<string, string>,
	) {}

	static async load(bytes: Uint8Array | ArrayBuffer): Promise<Package> {
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(bytes);
		} catch {
			throw new Error("The file is not a valid .xlsx workbook.");
		}
		const texts = new Map<string, string>();
		const xmlFiles = Object.values(zip.files).filter((f) => !f.dir && /\.(xml|rels)$/i.test(f.name));
		await Promise.all(xmlFiles.map(async (f) => texts.set(f.name, await f.async("string"))));
		return new Package(zip, texts);
	}

	xml(path: string): Document | undefined {
		let doc = this.docs.get(path);
		if (!doc) {
			const text = this.texts.get(path);
			if (text === undefined) return undefined;
			doc = parseXml(text);
			this.docs.set(path, doc);
		}
		return doc;
	}

	markDirty(path: string) {
		this.dirty.add(path);
	}

	/** Relationships of a part; pass "" for the package-level ones. */
	rels(partPath: string): Relationship[] {
		const doc = this.xml(relsPathOf(partPath));
		if (!doc) return [];
		const baseDir = partPath.includes("/") ? partPath.slice(0, partPath.lastIndexOf("/")) : "";
		return elements(doc.documentElement, "Relationship")
			.filter((r) => r.getAttribute("TargetMode") !== "External")
			.map((r) => ({
				id: r.getAttribute("Id") ?? "",
				type: r.getAttribute("Type") ?? "",
				target: resolveTarget(baseDir, r.getAttribute("Target") ?? ""),
			}));
	}

	/** Deletes a part together with its relationship from `ownerPath` and its content-type entry. */
	removePart(path: string, ownerPath: string) {
		if (!this.zip.file(path)) return;
		this.zip.remove(path);
		this.texts.delete(path);
		this.docs.delete(path);
		this.dirty.delete(path);

		const relsPath = relsPathOf(ownerPath);
		const relsDoc = this.xml(relsPath);
		const ownerRel = this.rels(ownerPath).find((r) => r.target === path);
		if (relsDoc && ownerRel) {
			for (const r of elements(relsDoc.documentElement, "Relationship")) {
				if (r.getAttribute("Id") === ownerRel.id) relsDoc.documentElement.removeChild(r);
			}
			this.markDirty(relsPath);
		}

		const types = this.xml(CONTENT_TYPES);
		if (types) {
			for (const o of elements(types.documentElement, "Override")) {
				if (o.getAttribute("PartName") === `/${path}`) types.documentElement.removeChild(o);
			}
			this.markDirty(CONTENT_TYPES);
		}
	}

	async toBytes(): Promise<Uint8Array> {
		for (const path of this.dirty) {
			const doc = this.docs.get(path);
			if (doc) this.zip.file(path, serializeXml(doc));
		}
		this.dirty.clear();
		return this.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
	}
}

function relsPathOf(partPath: string): string {
	if (!partPath) return "_rels/.rels";
	const slash = partPath.lastIndexOf("/");
	return `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
}

function resolveTarget(baseDir: string, target: string): string {
	const parts = target.startsWith("/") ? [] : baseDir.split("/").filter(Boolean);
	for (const segment of target.split("/")) {
		if (segment === "..") parts.pop();
		else if (segment && segment !== ".") parts.push(segment);
	}
	return parts.join("/");
}
