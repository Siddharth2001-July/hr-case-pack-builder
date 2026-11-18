/**
 * Document store abstraction for the case pack builder.
 *
 * The POC runs against `MockSharePointStore`, an in-memory stand-in for a SharePoint Online
 * document library. In production it is swapped for a Graph-backed implementation of the
 * same interface. Each method notes the Microsoft Graph call it maps to.
 *
 * Nothing in this file writes to disk: files live in memory (ArrayBuffer) for the life of the tab,
 * which mirrors the "no local caching / no OneDrive sync" requirement.
 */

export interface StoreItem {
  /** Path relative to the library root, e.g. "HR-2026-0147/01 Referral letter.docx". */
  path: string;
  name: string;
  size: number;
  isFolder: boolean;
  modified: Date;
  modifiedBy: string;
}

export interface CaseFolder {
  name: string;
  title: string;
}

export interface DocumentStore {
  readonly siteName: string;
  readonly libraryName: string;
  readonly user: { displayName: string; email: string };

  /** Graph: GET /sites/{site-id}/drives/{drive-id}/root/children */
  listCaseFolders(): Promise<CaseFolder[]>;

  /** Graph: GET /sites/{site-id}/drives/{drive-id}/root:/{folderPath}:/children */
  listFolder(folderPath: string): Promise<StoreItem[]>;

  /** Graph: GET /sites/{site-id}/drives/{drive-id}/root:/{path}:/content  (read into memory, never saved to disk) */
  download(path: string): Promise<ArrayBuffer>;

  /**
   * Graph: PUT /sites/{site-id}/drives/{drive-id}/root:/{path}:/content for files under 4 MB,
   * otherwise POST .../root:/{path}:/createUploadSession and upload in chunks.
   * Missing parent folders are created by Graph automatically on upload.
   */
  upload(path: string, data: ArrayBuffer, contentType: string): Promise<StoreItem>;
}

interface LibraryIndex {
  site: string;
  library: string;
  folders: { name: string; title: string; files: { name: string; path: string; size: number }[] }[];
}

interface Entry {
  item: StoreItem;
  data?: ArrayBuffer;
  sourceUrl?: string;
}

const delay = (ms: number) => new Promise((r) => setTimeout(r, ms));

/**
 * In-memory mock of a SharePoint document library, seeded from the generated sample case files.
 * Adds a small artificial latency so the UI behaves like it's talking to a remote service.
 */
export class MockSharePointStore implements DocumentStore {
  siteName = "";
  libraryName = "";
  readonly user = { displayName: "Alex Morgan (HR Adviser)", email: "alex.morgan@example.org" };

  private entries = new Map<string, Entry>();
  private folders: CaseFolder[] = [];

  static async create(indexUrl = "samples/library.json"): Promise<MockSharePointStore> {
    const store = new MockSharePointStore();
    const index: LibraryIndex = await fetch(indexUrl).then((r) => r.json());
    store.siteName = index.site;
    store.libraryName = index.library;
    const seeded = new Date(Date.now() - 1000 * 60 * 60 * 24 * 3);
    for (const f of index.folders) {
      store.folders.push({ name: f.name, title: f.title });
      store.addFolder(f.name);
      for (const file of f.files) {
        const path = `${f.name}/${file.name}`;
        store.entries.set(path, {
          item: { path, name: file.name, size: file.size, isFolder: false, modified: seeded, modifiedBy: "Helen Carver" },
          sourceUrl: file.path,
        });
      }
    }
    return store;
  }

  async listCaseFolders(): Promise<CaseFolder[]> {
    await delay(120);
    return [...this.folders];
  }

  async listFolder(folderPath: string): Promise<StoreItem[]> {
    await delay(120);
    const prefix = folderPath.replace(/\/$/, "") + "/";
    return [...this.entries.values()]
      .map((e) => e.item)
      .filter((i) => i.path.startsWith(prefix) && !i.path.slice(prefix.length).includes("/"))
      .sort((a, b) => Number(b.isFolder) - Number(a.isFolder) || a.name.localeCompare(b.name));
  }

  async download(path: string): Promise<ArrayBuffer> {
    const e = this.entries.get(path);
    if (!e || e.item.isFolder) throw new Error(`Not found: ${path}`);
    if (!e.data) {
      // Simulates the Graph content download. Fetched once, then held in memory only.
      e.data = await fetch(encodeURI(e.sourceUrl!)).then((r) => {
        if (!r.ok) throw new Error(`Download failed (${r.status}): ${path}`);
        return r.arrayBuffer();
      });
    } else {
      await delay(80);
    }
    return e.data!.slice(0);
  }

  async upload(path: string, data: ArrayBuffer, _contentType: string): Promise<StoreItem> {
    await delay(250);
    const parts = path.split("/");
    for (let i = 1; i < parts.length; i++) this.addFolder(parts.slice(0, i).join("/"));
    const item: StoreItem = {
      path,
      name: parts[parts.length - 1],
      size: data.byteLength,
      isFolder: false,
      modified: new Date(),
      modifiedBy: this.user.displayName,
    };
    this.entries.set(path, { item, data: data.slice(0) });
    return item;
  }

  private addFolder(path: string) {
    if (this.entries.has(path)) return;
    const name = path.split("/").pop()!;
    this.entries.set(path, {
      item: { path, name, size: 0, isFolder: true, modified: new Date(), modifiedBy: this.user.displayName },
    });
  }
}
