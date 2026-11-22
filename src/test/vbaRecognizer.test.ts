import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { getLogicalLinePrefix, recognizeVbaContext } from "../intellisense/vbaRecognizer";

describe("recognizeVbaContext", () => {
    it("recognizes the forms context after 'Forms!'", () => {
        assert.deepEqual(recognizeVbaContext("Forms!"), { kind: "forms", partial: "" });
        assert.deepEqual(recognizeVbaContext("    x = forms!frm"), { kind: "forms", partial: "frm" });
        assert.deepEqual(recognizeVbaContext("Debug.Print \"a\" & Forms!"), { kind: "forms", partial: "" });
        assert.deepEqual(recognizeVbaContext("Forms![Order"), { kind: "forms", partial: "[Order" });
        assert.deepEqual(recognizeVbaContext("Forms![Customer's"), { kind: "forms", partial: "[Customer's" });
        assert.deepEqual(recognizeVbaContext("Remote = Forms!"), { kind: "forms", partial: "" });
    });

    it("recognizes the controls context and extracts the form name", () => {
        assert.deepEqual(
            recognizeVbaContext("Forms!frmCustomers!"),
            { kind: "formControls", formName: "frmCustomers", partial: "" }
        );
        assert.deepEqual(
            recognizeVbaContext("x = Forms!frmCustomers!Cust"),
            { kind: "formControls", formName: "frmCustomers", partial: "Cust" }
        );
        assert.deepEqual(
            recognizeVbaContext("Forms![Order Entry]!"),
            { kind: "formControls", formName: "Order Entry", partial: "" }
        );
        assert.deepEqual(
            recognizeVbaContext("Forms![Customer's \"Orders\"]!"),
            { kind: "formControls", formName: "Customer's \"Orders\"", partial: "" }
        );
    });

    it("does not trigger inside comments", () => {
        assert.equal(recognizeVbaContext("' Forms!"), undefined);
        assert.equal(recognizeVbaContext("x = 1 ' Forms!frmCustomers!"), undefined);
        assert.equal(recognizeVbaContext("Rem Forms!"), undefined);
        assert.equal(recognizeVbaContext("x = 1: Rem Forms!"), undefined);
    });

    it("does not trigger inside string literals", () => {
        assert.equal(recognizeVbaContext("Debug.Print \"Forms!"), undefined);
        assert.equal(recognizeVbaContext("Debug.Print \"say \"\"hi\"\" Forms!"), undefined);
    });

    it("ignores unrelated or unsupported syntax", () => {
        assert.equal(recognizeVbaContext("Dim x As Long"), undefined);
        assert.equal(recognizeVbaContext("Application.Forms!"), undefined);
        assert.equal(recognizeVbaContext("MyForms!"), undefined);
        assert.equal(recognizeVbaContext("Forms!frmCustomers!sub!"), undefined);
        assert.equal(recognizeVbaContext("Forms(\"frmCustomers\")!"), undefined);
    });
});

describe("getLogicalLinePrefix", () => {
    it("joins lines ended with a continuation character", () => {
        const lines = ["x = Foo(1, _", "    Forms!", "unrelated"];
        const prefix = getLogicalLinePrefix((index) => lines[index], 1, lines[1].length);
        assert.deepEqual(recognizeVbaContext(prefix), { kind: "forms", partial: "" });
    });

    it("does not join lines that are not continued", () => {
        const lines = ["' note Forms!", "x = 1"];
        assert.equal(getLogicalLinePrefix((index) => lines[index], 1, lines[1].length), "x = 1");
    });
});
