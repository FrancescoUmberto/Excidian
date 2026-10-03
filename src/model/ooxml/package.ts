import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from "fflate";
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
 * dirty are rewritten on save; the content of every other part is kept as is.
 */
export class Package {
	private docs = new Map<string, Document>();
	private dirty = new Set<string>();

	/** Entries in their original order (Excel expects [Content_Types].xml first). */
	private constructor(private entries: Record<string, Uint8Array>) {}

	static async load(bytes: Uint8Array | ArrayBuffer): Promise<Package> {
		try {
			return new Package(unzipSync(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)));
		} catch {
			throw new Error("The file is not a valid .xlsx workbook.");
		}
	}

	xml(path: string): Document | undefined {
		let doc = this.docs.get(path);
		if (!doc) {
			const data = this.entries[path];
			if (data === undefined || path.endsWith("/")) return undefined;
			doc = parseXml(strFromU8(data));
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
		if (!(path in this.entries)) return;
		delete this.entries[path];
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
			if (doc) this.entries[path] = strToU8(serializeXml(doc));
		}
		this.dirty.clear();
		const files: Zippable = {};
		for (const [path, data] of Object.entries(this.entries)) files[path] = [data, { level: path.endsWith("/") ? 0 : 6 }];
		return zipSync(files);
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
