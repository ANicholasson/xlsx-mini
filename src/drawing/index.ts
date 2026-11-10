import { XlsxMiniError, describeValue } from "../errors/index.js";
import { crc32 } from "../opc/index.js";
import { XML_DECLARATION } from "../xml/index.js";

/** Excel's 96 dpi: one pixel in English Metric Units. */
const EMU_PER_PIXEL = 9525;
/** The gap between a picture and its cell's top-left corner, in pixels. */
export const IMAGE_INSET = 2;
const MAX_IMAGE_PIXELS = 100_000;
const RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const IMAGE_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const DRAWING_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const DRAWING_MAIN_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const OFFICE_RELATIONSHIPS_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** The relationship type from a worksheet to its drawing part. */
export const DRAWING_RELATIONSHIP = `${OFFICE_RELATIONSHIPS_NAMESPACE}/drawing`;

/** The path of a drawing part and of a media part, from their one-based numbers. */
export const drawingPath = (n: number): string => `xl/drawings/drawing${n}.xml`;
export const drawingRelationshipsPath = (n: number): string =>
  `xl/drawings/_rels/drawing${n}.xml.rels`;
const mediaPath = (n: number): string => `xl/media/image${n}.png`;

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

/** A registered image: its one-based media number and its pixel size from the header. */
interface Media {
  readonly id: number;
  readonly width: number;
  readonly height: number;
}

/** A media part to write: `xl/media/imageN.png` and the bytes as given. */
interface MediaPart {
  readonly path: string;
  readonly data: Uint8Array;
}

/**
 * The workbook's images, one media part per distinct image however many cells show it, numbered
 * across sheets. Identical bytes are found by length and CRC-32 and confirmed byte for byte.
 */
export interface MediaTable {
  /** Registers PNG bytes for the cell at `path`, or throws `INVALID_IMAGE` when they are not a PNG. */
  add(bytes: unknown, path: string): Media;
  /** The media parts registered since the last call, in order, so each sheet's follow the sheet. */
  takeParts(): MediaPart[];
}

/** The pixel size in a PNG's IHDR chunk, or a reason the bytes are not a PNG. */
function pngSize(bytes: Uint8Array): { width: number; height: number } | string {
  if (bytes.length < 24) return "the bytes are too short to be a PNG";
  for (let i = 0; i < PNG_SIGNATURE.length; i++) {
    if (bytes[i] !== PNG_SIGNATURE[i]) return "the bytes do not start with the PNG signature";
  }
  if (bytes[12] !== 0x49 || bytes[13] !== 0x48 || bytes[14] !== 0x44 || bytes[15] !== 0x52) {
    return "the PNG does not start with its IHDR chunk";
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const width = view.getUint32(16);
  const height = view.getUint32(20);
  if (width === 0 || height === 0 || width > MAX_IMAGE_PIXELS || height > MAX_IMAGE_PIXELS) {
    return `the PNG header gives an impossible size of ${width} by ${height} pixels`;
  }
  return { width, height };
}

/** Byte-for-byte equality of two arrays the key has already matched by length. */
function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

export function createMediaTable(): MediaTable {
  /** Registered images by `length:crc32`; a list, since different bytes can share both. */
  const byKey = new Map<string, { bytes: Uint8Array; media: Media }[]>();
  let count = 0;
  let pending: MediaPart[] = [];
  return {
    add(bytes, path) {
      if (!(bytes instanceof Uint8Array)) {
        throw new XlsxMiniError(
          "INVALID_IMAGE",
          `${path}.image`,
          `An image is the bytes of a PNG as a Uint8Array, not ${describeValue(bytes)}`,
        );
      }
      const size = pngSize(bytes);
      if (typeof size === "string") {
        throw new XlsxMiniError("INVALID_IMAGE", `${path}.image`, `Not a PNG image: ${size}`);
      }
      const key = `${bytes.length}:${crc32(bytes)}`;
      let known = byKey.get(key);
      if (known === undefined) {
        known = [];
        byKey.set(key, known);
      }
      for (const entry of known) if (sameBytes(entry.bytes, bytes)) return entry.media;
      count++;
      const media: Media = { id: count, ...size };
      known.push({ bytes, media });
      pending.push({ path: mediaPath(count), data: bytes });
      return media;
    },
    takeParts() {
      const parts = pending;
      pending = [];
      return parts;
    },
  };
}

/** The size a picture is drawn at, in pixels. */
interface PictureSize {
  readonly width: number;
  readonly height: number;
}

interface Anchor {
  readonly column: number;
  readonly rowIndex: number;
  /** The one-based relationship number in the drawing's relationships part. */
  readonly rId: number;
  readonly size: PictureSize;
}

/**
 * One sheet's drawing part: the anchors recorded as its rows were written, each pointing at a
 * media part through the drawing's own relationships.
 */
export interface SheetDrawing {
  /**
   * Anchors the registered image at a zero-based column and row, at the size requested or the
   * image's own, one given dimension scaling the other. Returns the size drawn, so the row and the
   * column can make room for it.
   */
  add(
    column: number,
    rowIndex: number,
    media: Media,
    width: number | undefined,
    height: number | undefined,
    path: string,
  ): PictureSize;
  /** True once any image has been anchored, so the sheet knows to carry a drawing. */
  readonly hasImages: boolean;
  /** `xl/drawings/drawingN.xml`. */
  render(): string;
  /** `xl/drawings/_rels/drawingN.xml.rels`: one relationship per distinct media part the sheet shows. */
  relationships(): string;
}

/** A requested dimension in pixels: a positive finite number, or an error at the property. */
function dimension(value: unknown, path: string, name: string): number | undefined {
  if (value === undefined) return undefined;
  if (
    typeof value !== "number" ||
    !(value > 0) ||
    !Number.isFinite(value) ||
    value > MAX_IMAGE_PIXELS
  ) {
    throw new XlsxMiniError(
      "INVALID_IMAGE",
      `${path}.${name}`,
      `An image ${name} is a number of pixels from above 0 to ${MAX_IMAGE_PIXELS}, not ${describeValue(value)}`,
    );
  }
  return value;
}

export function createSheetDrawing(): SheetDrawing {
  const anchors: Anchor[] = [];
  /** Relationship number by media id. */
  const relationships = new Map<number, number>();
  const self: SheetDrawing = {
    add(column, rowIndex, media, width, height, path) {
      const w = dimension(width, path, "width");
      const h = dimension(height, path, "height");
      const size: PictureSize =
        w === undefined && h === undefined
          ? { width: media.width, height: media.height }
          : {
              width: w ?? (h! * media.width) / media.height,
              height: h ?? (w! * media.height) / media.width,
            };
      let rId = relationships.get(media.id);
      if (rId === undefined) {
        rId = relationships.size + 1;
        relationships.set(media.id, rId);
      }
      anchors.push({ column, rowIndex, rId, size });
      return size;
    },
    get hasImages() {
      return anchors.length > 0;
    },
    render() {
      const inset = IMAGE_INSET * EMU_PER_PIXEL;
      let xml = `${XML_DECLARATION}<xdr:wsDr xmlns:xdr="${DRAWING_NAMESPACE}" xmlns:a="${DRAWING_MAIN_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}">`;
      for (let i = 0; i < anchors.length; i++) {
        const anchor = anchors[i]!;
        const cx = Math.round(anchor.size.width * EMU_PER_PIXEL);
        const cy = Math.round(anchor.size.height * EMU_PER_PIXEL);
        xml +=
          "<xdr:oneCellAnchor>" +
          `<xdr:from><xdr:col>${anchor.column}</xdr:col><xdr:colOff>${inset}</xdr:colOff><xdr:row>${anchor.rowIndex}</xdr:row><xdr:rowOff>${inset}</xdr:rowOff></xdr:from>` +
          `<xdr:ext cx="${cx}" cy="${cy}"/>` +
          "<xdr:pic>" +
          `<xdr:nvPicPr><xdr:cNvPr id="${i + 2}" name="Picture ${i + 1}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
          `<xdr:blipFill><a:blip r:embed="rId${anchor.rId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
          `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>` +
          "</xdr:pic>" +
          "<xdr:clientData/>" +
          "</xdr:oneCellAnchor>";
      }
      return `${xml}</xdr:wsDr>`;
    },
    relationships() {
      let xml = `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">`;
      for (const [mediaId, rId] of relationships) {
        xml += `<Relationship Id="rId${rId}" Type="${IMAGE_RELATIONSHIP}" Target="../media/image${mediaId}.png"/>`;
      }
      return `${xml}</Relationships>`;
    },
  };
  return self;
}
