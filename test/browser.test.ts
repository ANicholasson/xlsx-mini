import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { XLSX_MIME_TYPE, downloadXlsx } from "../src/browser.js";
import { XLSX_MIME_TYPE as CORE_MIME_TYPE, XlsxMiniError } from "../src/index.js";

const bytes = new TextEncoder().encode("PK\u0003\u0004 stands in for a workbook");

/** The parts of a document the download touches, recording instead of rendering. */
function fakeDocument() {
  const anchor = { href: "", download: "", click: vi.fn(), remove: vi.fn() };
  const body = { append: vi.fn() };
  const document = { createElement: vi.fn(() => anchor), body };
  return { document, anchor, body };
}

function caught(fn: () => void): unknown {
  try {
    fn();
  } catch (error) {
    return error;
  }
  return undefined;
}

describe("downloadXlsx", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    vi.useRealTimers();
  });

  it("imports without a document and reports NO_DOM when called outside a browser", () => {
    expect(typeof document).toBe("undefined");
    const error = caught(() => downloadXlsx(bytes, "report"));
    expect(error).toBeInstanceOf(XlsxMiniError);
    expect(error).toMatchObject({ code: "NO_DOM", path: "downloadXlsx" });
  });

  it("saves the bytes through an anchor click, then revokes the object URL later", () => {
    vi.useFakeTimers();
    const { document, anchor, body } = fakeDocument();
    vi.stubGlobal("document", document);
    const createObjectURL = vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:xlsx-mini/1");
    const revokeObjectURL = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => undefined);

    downloadXlsx(bytes, "orders");

    const blob = createObjectURL.mock.calls[0]![0] as Blob;
    expect(blob.type).toBe(XLSX_MIME_TYPE);
    expect(blob.size).toBe(bytes.byteLength);
    expect(document.createElement).toHaveBeenCalledWith("a");
    expect(anchor.href).toBe("blob:xlsx-mini/1");
    expect(anchor.download).toBe("orders.xlsx");
    expect(body.append).toHaveBeenCalledWith(anchor);
    expect(anchor.click).toHaveBeenCalledOnce();
    expect(anchor.remove).toHaveBeenCalledOnce();

    expect(revokeObjectURL).not.toHaveBeenCalled();
    vi.advanceTimersByTime(40_000);
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:xlsx-mini/1");
  });

  it("keeps a filename that already carries the extension, whatever its case", () => {
    const { document, anchor } = fakeDocument();
    vi.stubGlobal("document", document);
    vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:xlsx-mini/2");
    vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => undefined);

    downloadXlsx(bytes, "Report.XLSX");

    expect(anchor.download).toBe("Report.XLSX");
  });

  it("exports the same media type as the core entry", () => {
    expect(XLSX_MIME_TYPE).toBe(CORE_MIME_TYPE);
  });
});
