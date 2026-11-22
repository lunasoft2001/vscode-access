import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import type { AccessConnection, AccessControlInfo, AccessObjectInfo, AccessTableFieldInfo } from "../models/types";
import { AccessMetadataService, BackgroundMetadataBlockedError } from "../services/accessMetadataService";
import type { AccessMetadataSource } from "../services/accessMetadataService";

const connectionA: AccessConnection = { id: "a", name: "Sales", dbPath: "C:\\db\\sales.accdb" };
const connectionB: AccessConnection = { id: "b", name: "HR", dbPath: "C:\\db\\hr.accdb" };

class FakeSource implements AccessMetadataSource {
    readonly calls: string[] = [];
    gate: Promise<void> | undefined;
    failNext = false;
    suffix = "";

    async listObjects(connection: AccessConnection, objectType: string): Promise<AccessObjectInfo[]> {
        await this.enter(`list:${objectType}:${connection.id}`);
        const names = objectType === "form" ? ["frmCustomers", "frmOrders"] : ["Customers", "Orders"];
        return names.map((name) => ({ name: name + this.suffix } as AccessObjectInfo));
    }

    async getControls(connection: AccessConnection, _type: "form" | "report", name: string): Promise<AccessControlInfo[]> {
        await this.enter(`controls:${connection.id}:${name}`);
        return [{ name: "CustomerName", type_name: "TextBox" } as AccessControlInfo];
    }

    async getTableFields(connection: AccessConnection, name: string): Promise<AccessTableFieldInfo[]> {
        await this.enter(`fields:${connection.id}:${name}`);
        return [{ name: "OrderID", type: "LONG", required: true } as AccessTableFieldInfo];
    }

    private async enter(call: string): Promise<void> {
        this.calls.push(call);
        await this.gate;
        if (this.failNext) {
            this.failNext = false;
            throw new Error("boom");
        }
    }
}

describe("AccessMetadataService", () => {
    it("loads on first request and serves the cache afterwards", async () => {
        const source = new FakeSource();
        const service = new AccessMetadataService(source);

        assert.deepEqual((await service.getForms(connectionA)).map((f) => f.name), ["frmCustomers", "frmOrders"]);
        await service.getForms(connectionA);
        await service.getTables(connectionA);
        await service.getTables(connectionA);
        await service.getFormControls(connectionA, "frmCustomers");
        await service.getFormControls(connectionA, "FRMCUSTOMERS");
        await service.getTableFields(connectionA, "Orders");
        await service.getTableFields(connectionA, "orders");

        assert.deepEqual(source.calls, [
            "list:form:a", "list:table:a", "controls:a:frmCustomers", "fields:a:Orders"
        ]);
    });

    it("exposes plain data without extra properties", async () => {
        const service = new AccessMetadataService(new FakeSource());
        assert.deepEqual(await service.getFormControls(connectionA, "frmCustomers"), [{ name: "CustomerName", type: "TextBox" }]);
        assert.deepEqual(await service.getTableFields(connectionA, "Orders"), [
            { name: "OrderID", type: "LONG", size: undefined, required: true }
        ]);
    });

    it("deduplicates simultaneous requests for the same item", async () => {
        const source = new FakeSource();
        let release!: () => void;
        source.gate = new Promise<void>((resolve) => { release = resolve; });
        const service = new AccessMetadataService(source);

        const pending = [service.getForms(connectionA), service.getForms(connectionA), service.getForms(connectionA)];
        release();
        const results = await Promise.all(pending);

        assert.equal(source.calls.length, 1);
        assert.equal(results[0], results[1]);
    });

    it("reloads after invalidation", async () => {
        const source = new FakeSource();
        const service = new AccessMetadataService(source);

        await service.getForms(connectionA);
        await service.getForms(connectionB);
        service.invalidate(connectionA.id);
        source.suffix = "2";

        assert.equal((await service.getForms(connectionA))[0].name, "frmCustomers2");
        assert.equal((await service.getForms(connectionB))[0].name, "frmCustomers");
        assert.deepEqual(source.calls, ["list:form:a", "list:form:b", "list:form:a"]);

        service.invalidateAll();
        await service.getForms(connectionB);
        assert.equal(source.calls.length, 4);
    });

    it("does not let an in-flight load repopulate the cache after invalidation", async () => {
        const source = new FakeSource();
        let release!: () => void;
        source.gate = new Promise<void>((resolve) => { release = resolve; });
        const service = new AccessMetadataService(source);

        const stale = service.getForms(connectionA);
        service.invalidate(connectionA.id);
        source.gate = undefined;
        source.suffix = "-new";
        const fresh = await service.getForms(connectionA);
        release();
        assert.deepEqual(await stale, [], "invalidated requests must not return stale metadata");

        assert.equal(fresh[0].name, "frmCustomers-new");
        assert.equal((await service.getForms(connectionA))[0].name, "frmCustomers-new");
        assert.equal(source.calls.length, 2);
    });

    it("does not share metadata between connections", async () => {
        const source = new FakeSource();
        const service = new AccessMetadataService(source);

        await service.getTableFields(connectionA, "Orders");
        await service.getTableFields(connectionB, "Orders");

        assert.deepEqual(source.calls, ["fields:a:Orders", "fields:b:Orders"]);
    });

    it("discards a snapshot whose connection now points to another database", async () => {
        const source = new FakeSource();
        const service = new AccessMetadataService(source);

        await service.getForms(connectionA);
        await service.getForms({ ...connectionA, dbPath: "C:\\db\\other.accdb" });

        assert.equal(source.calls.length, 2);
    });

    it("resolves failures to an empty list, caches nothing and backs off before retrying", async () => {
        const source = new FakeSource();
        let now = 1_000;
        const diagnostics: string[] = [];
        const service = new AccessMetadataService(source, { now: () => now, onDiagnostic: (m) => diagnostics.push(m) });

        source.failNext = true;
        assert.deepEqual(await service.getForms(connectionA), []);
        assert.equal(diagnostics.length, 1);

        assert.deepEqual(await service.getForms(connectionA), []);
        assert.equal(source.calls.length, 1, "no request storm during the cooldown");

        now += 20_000;
        assert.equal((await service.getForms(connectionA)).length, 2);
        assert.equal(source.calls.length, 2);
    });

    it("never calls Access when loading is not allowed, but still serves cached data", async () => {
        const source = new FakeSource();
        let allowed = true;
        const blocked: string[] = [];
        const service = new AccessMetadataService(source, {
            canLoad: () => allowed,
            onLoadBlocked: (connection) => blocked.push(connection.name)
        });

        await service.getForms(connectionA);
        allowed = false;

        assert.equal((await service.getForms(connectionA)).length, 2);
        assert.deepEqual(await service.getTables(connectionA), []);
        assert.deepEqual(blocked, ["Sales"]);
        assert.equal(source.calls.length, 1);

        allowed = true;
        assert.equal((await service.getTables(connectionA)).length, 2);
    });

    it("reports a dispatch-time block without cooling down subsequent requests", async () => {
        const source = new FakeSource();
        const listObjects = source.listObjects.bind(source);
        let blocked = true;
        source.listObjects = async (...args) => {
            if (blocked) {
                blocked = false;
                throw new BackgroundMetadataBlockedError();
            }
            return listObjects(...args);
        };
        const notices: string[] = [];
        const service = new AccessMetadataService(source, { onLoadBlocked: (item) => notices.push(item.name) });
        assert.deepEqual(await service.getForms(connectionA), []);
        assert.deepEqual(notices, ["Sales"]);
        assert.equal((await service.getForms(connectionA)).length, 2);
    });
});
