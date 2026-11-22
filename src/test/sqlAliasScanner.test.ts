import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { findSourceBindings, resolveSqlQualifier } from "../intellisense/sqlAliasScanner";

/** Splits a "|" cursor marker out of the sample text. */
function withCursor(sample: string): { text: string; offset: number } {
    const offset = sample.indexOf("|");
    assert.notEqual(offset, -1, "sample needs a | cursor marker");
    return { text: sample.slice(0, offset) + sample.slice(offset + 1), offset };
}

function resolve(sample: string) {
    const { text, offset } = withCursor(sample);
    return resolveSqlQualifier(text, offset);
}

describe("resolveSqlQualifier", () => {
    it("binds an alias to its table (FROM Table alias)", () => {
        const { text, offset } = withCursor("SELECT o.|\nFROM Orders o");
        assert.deepEqual(resolveSqlQualifier(text, offset), {
            qualifier: "o",
            source: "Orders",
            partial: "",
            partialStart: offset
        });
        assert.deepEqual(findSourceBindings(text, offset), [{ alias: "o", source: "Orders" }]);
    });

    it("binds an alias declared with AS", () => {
        assert.equal(resolve("SELECT o.|\nFROM Orders AS o")?.source, "Orders");
    });

    it("binds aliases of a simple JOIN", () => {
        const sql = "SELECT |.\nFROM Customers c\nINNER JOIN Orders o ON c.CustomerID = o.CustomerID";
        assert.equal(resolve(sql.replace("|.", "o.|"))?.source, "Orders");
        assert.equal(resolve(sql.replace("|.", "c.|"))?.source, "Customers");
    });

    it("captures what was already typed after the dot", () => {
        const { text, offset } = withCursor("SELECT o.Ord|\nFROM Orders o");
        const context = resolveSqlQualifier(text, offset);
        assert.equal(context?.partial, "Ord");
        assert.equal(context?.partialStart, offset - 3);
        assert.equal(resolve("SELECT o.[Ord|\nFROM Orders o")?.partial, "[Ord");
    });

    it("returns nothing for an unknown alias", () => {
        const { text, offset } = withCursor("SELECT x.|\nFROM Orders o");
        assert.equal(resolveSqlQualifier(text, offset), undefined);
        assert.deepEqual(findSourceBindings(text, offset), [{ alias: "o", source: "Orders" }]);
    });

    it("supports bracketed names, unaliased tables, comma lists and Access join groups", () => {
        assert.equal(resolve("SELECT od.|\nFROM [Order Details] od")?.source, "Order Details");
        assert.equal(resolve("SELECT [od].|\nFROM [Order Details] AS [od]")?.source, "Order Details");
        assert.equal(resolve("SELECT Orders.|\nFROM Orders WHERE 1 = 1")?.source, "Orders");
        assert.equal(resolve("SELECT c.|\nFROM Orders o, Customers c")?.source, "Customers");
        assert.equal(resolve("SELECT o.|\nFROM (Customers c INNER JOIN Orders o ON c.ID = o.CustomerID)")?.source, "Orders");
    });

    it("does not let reserved words become aliases", () => {
        assert.equal(resolve("SELECT Orders.|\nFROM Orders\nWHERE Orders.ID > 1")?.source, "Orders");
        assert.deepEqual(findSourceBindings("SELECT * FROM Orders ORDER BY ID", 0), [{ alias: "Orders", source: "Orders" }]);
    });

    it("only looks at the statement that contains the cursor", () => {
        assert.equal(resolve("SELECT a.x FROM Customers a; SELECT o.|\nFROM Orders o")?.source, "Orders");
        assert.equal(resolve("SELECT a.x FROM Customers a; SELECT a.|\nFROM Orders o"), undefined);
    });

    it("produces no suggestions for unsupported syntax", () => {
        assert.equal(resolve("SELECT o.|\nFROM (SELECT * FROM Orders) o"), undefined);
        assert.equal(resolve("SELECT o.| FROM Orders o UNION SELECT o.ID FROM Orders o"), undefined);
        assert.equal(resolve("TRANSFORM Sum(o.|) SELECT o.ID FROM Orders o PIVOT o.Year"), undefined);
        assert.equal(resolve("UPDATE Orders o SET o.|"), undefined);
        assert.equal(resolve("SELECT o.|\nFROM Orders o, Customers o"), undefined);
        assert.equal(resolve("SELECT o.|\nFROM dbo.Orders o"), undefined);
        assert.equal(resolve("SELECT Orders.|\nFROM Orders IN 'other.accdb'"), undefined);
        assert.equal(resolve("SELECT o.| FROM Orders AS o IN 'other.accdb'"), undefined);
        assert.equal(resolve("SELECT o.| FROM Orders o IN 'other.accdb'"), undefined);
        assert.equal(resolve("SELECT db.o.|\nFROM Orders o"), undefined);
    });

    it("does not trigger inside strings, comments or without a dot", () => {
        assert.equal(resolve("SELECT 'o.|' FROM Orders o"), undefined);
        assert.equal(resolve("SELECT * FROM Orders o -- o.|"), undefined);
        assert.equal(resolve("SELECT * /* o.| */ FROM Orders o"), undefined);
        assert.equal(resolve("SELECT o|\nFROM Orders o"), undefined);
        assert.equal(resolve("SELECT o. |\nFROM Orders o"), undefined);
    });
});
