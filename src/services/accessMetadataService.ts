import type { AccessConnection, AccessControlInfo, AccessObjectInfo, AccessTableFieldInfo } from "../models/types";

export interface FormInfo {
    name: string;
}

export interface ControlInfo {
    name: string;
    type?: string;
}

export interface TableInfo {
    name: string;
}

export interface FieldInfo {
    name: string;
    type?: string;
    size?: number;
    required?: boolean;
}

/** The subset of McpAccessClient the service needs; all calls are read-only metadata queries. */
export interface AccessMetadataSource {
    listObjects(connection: AccessConnection, objectType: string): Promise<AccessObjectInfo[]>;
    getControls(connection: AccessConnection, objectType: "form" | "report", objectName: string): Promise<AccessControlInfo[]>;
    getTableFields(connection: AccessConnection, tableName: string): Promise<AccessTableFieldInfo[]>;
}

export interface AccessMetadataServiceOptions {
    /** Returning false makes a cache miss resolve to [] without calling Access (cached data is still served). */
    canLoad?: (connection: AccessConnection) => boolean;
    /** Called when canLoad refused a load, so the UI can explain why nothing was suggested. */
    onLoadBlocked?: (connection: AccessConnection) => void;
    onDiagnostic?: (message: string) => void;
    now?: () => number;
}

const FAILURE_COOLDOWN_MS = 15_000;

interface Slot {
    value?: readonly unknown[];
    pending?: Promise<readonly unknown[]>;
    failedAt?: number;
}

interface Snapshot {
    dbPath: string;
    slots: Map<string, Slot>;
}

/**
 * In-memory, per-connection snapshot of plain metadata (names/types only) used by completion providers.
 * Everything loads lazily on first use, concurrent requests share one load, and results are kept until
 * invalidate()/invalidateAll(). Load failures resolve to [] so completion never throws.
 */
export class AccessMetadataService {
    private readonly snapshots = new Map<string, Snapshot>();
    private readonly now: () => number;

    constructor(
        private readonly source: AccessMetadataSource,
        private readonly options: AccessMetadataServiceOptions = {}
    ) {
        this.now = options.now ?? Date.now;
    }

    getForms(connection: AccessConnection): Promise<readonly FormInfo[]> {
        return this.load(connection, "forms", async () =>
            (await this.source.listObjects(connection, "form")).map((item) => ({ name: item.name }))
        );
    }

    getTables(connection: AccessConnection): Promise<readonly TableInfo[]> {
        return this.load(connection, "tables", async () =>
            (await this.source.listObjects(connection, "table")).map((item) => ({ name: item.name }))
        );
    }

    getFormControls(connection: AccessConnection, formName: string): Promise<readonly ControlInfo[]> {
        return this.load(connection, `controls:${formName.toLowerCase()}`, async () =>
            (await this.source.getControls(connection, "form", formName)).map((control) => ({
                name: control.name,
                type: control.type_name
            }))
        );
    }

    getTableFields(connection: AccessConnection, tableName: string): Promise<readonly FieldInfo[]> {
        return this.load(connection, `fields:${tableName.toLowerCase()}`, async () =>
            (await this.source.getTableFields(connection, tableName)).map((field) => ({
                name: field.name,
                type: field.type,
                size: field.size,
                required: field.required
            }))
        );
    }

    invalidate(connectionId: string): void {
        this.snapshots.delete(connectionId);
    }

    invalidateAll(): void {
        this.snapshots.clear();
    }

    private async load<T>(
        connection: AccessConnection,
        slotKey: string,
        loader: () => Promise<readonly T[]>
    ): Promise<readonly T[]> {
        const snapshot = this.getSnapshot(connection);
        let slot = snapshot.slots.get(slotKey);
        if (slot?.value) {
            return slot.value as readonly T[];
        }
        if (slot?.pending) {
            return slot.pending as Promise<readonly T[]>;
        }
        if (slot?.failedAt !== undefined && this.now() - slot.failedAt < FAILURE_COOLDOWN_MS) {
            return [];
        }
        if (this.options.canLoad && !this.options.canLoad(connection)) {
            this.options.onLoadBlocked?.(connection);
            return [];
        }

        const target: Slot = slot ?? {};
        snapshot.slots.set(slotKey, target);
        // A discarded snapshot keeps receiving its own in-flight result, but is never read again.
        target.pending = loader().then(
            (value) => {
                target.value = value;
                target.pending = undefined;
                target.failedAt = undefined;
                return this.snapshots.get(connection.id) === snapshot ? value : [];
            },
            (error) => {
                target.pending = undefined;
                target.failedAt = this.now();
                const message = error instanceof Error ? error.message : String(error);
                this.options.onDiagnostic?.(`Failed loading ${slotKey} for '${connection.name}': ${message}`);
                return [] as readonly T[];
            }
        );
        return target.pending as Promise<readonly T[]>;
    }

    private getSnapshot(connection: AccessConnection): Snapshot {
        const existing = this.snapshots.get(connection.id);
        if (existing && existing.dbPath === connection.dbPath) {
            return existing;
        }

        const created: Snapshot = { dbPath: connection.dbPath, slots: new Map() };
        this.snapshots.set(connection.id, created);
        return created;
    }
}
