import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

export type SpreadsheetRow = Record<string, string>;

export type PhotoArchive = {
  files: Record<string, File>;
  names: string[];
};

export type PhotoPlaceholder = {
  elementIndex: number;
  x: number;
  y: number;
  width: number;
  height: number;
  source: string;
};

export type SvgElementOption = {
  index: number;
  tag: string;
  id?: string;
  label: string;
  text: string;
  hasImageToken: boolean;
};

export type ParsedTemplate = {
  rawSvg: string;
  tokens: string[];
  photoPlaceholder?: PhotoPlaceholder;
  elements: SvgElementOption[];
};

export type SpreadsheetData = {
  headers: string[];
  rows: SpreadsheetRow[];
  embeddedPhotos?: Record<number, File>;
};

export type FieldMapping = Record<string, string>;

export type GenerationIssue = {
  row: number;
  field: string;
  message: string;
  person?: string;
};

export type GeneratedCard = {
  rowIndex: number;
  pngBlob: Blob;
  svgText: string;
};


/* =========================================================
   BASIC HELPERS
========================================================= */

const cleanToken = (value: string) =>
  value
    .trim()
    .replace(/^\{\{|\}\}$/g, '')
    .trim();

const baseName = (name: string) =>
  name
    .split('/')
    .pop()
    ?.replace(/\.[^.]+$/, '')
    .toLowerCase() ?? '';

const fileNameKey = (name: string) =>
  name.split('/').pop()?.toLowerCase() ?? '';

const escapeXml = (value: string) =>
  value.replace(
    /[<>&'"]/g,
    (character) =>
      ({
        '<': '&lt;',
        '>': '&gt;',
        '&': '&amp;',
        "'": '&apos;',
        '"': '&quot;',
      })[character] ?? character,
  );

function numeric(value: string | null, fallback: number) {
  const parsed = Number.parseFloat(value ?? '');
  return Number.isFinite(parsed) ? parsed : fallback;
}


/* =========================================================
   CANVA SVG SANITIZATION
========================================================= */

function sanitizeCanvaSvg(rawSvg: string): string {
  let cleaned = rawSvg;

  /*
   * Canva can split:
   *
   * {{image}}
   *
   * into multiple tspans.
   *
   * Example:
   *
   * <tspan>{{</tspan>
   * <tspan>image</tspan>
   * <tspan>}}</tspan>
   *
   * This merges those tspans when they belong to a token.
   */

  cleaned = cleaned.replace(
    /(<text[^>]*>)[\s\S]*?(<\/text>)/gi,
    (fullMatch) => {
      const textContentOnly = fullMatch.replace(/<[^>]+>/g, '');

      if (/\{\{[\s\S]*?\}\}/.test(textContentOnly)) {
        return fullMatch.replace(
          /<\/tspan>\s*<tspan[^>]*>/gi,
          '',
        );
      }

      return fullMatch;
    },
  );

  return cleaned;
}


/* =========================================================
   SVG ELEMENT LIST
========================================================= */

/**
 * Returns all SVG elements that the user can potentially select
 * as the image placeholder.
 *
 * IMPORTANT:
 * We do NOT automatically assume that <image> means photo.
 */
export function getSvgElements(rawSvg: string): SvgElementOption[] {
  const sanitizedSvg = sanitizeCanvaSvg(rawSvg);

  const doc = new DOMParser().parseFromString(
    sanitizedSvg,
    'image/svg+xml',
  );

  const elements = Array.from(doc.querySelectorAll('*'));

  return elements.map((element, index) => {
    const tag = element.tagName.toLowerCase();

    const id = element.getAttribute('id') ?? '';

    const text = element.textContent?.trim() ?? '';

    const hasImageToken =
      /\{\{\s*(image|photo)\s*\}\}/i.test(text);

    let label = '';

    if (id) {
      label = `#${id}`;
    } else if (text) {
      label = text.slice(0, 60);
    } else {
      label = `<${tag}>`;
    }

    return {
      index,
      tag,
      id: id || undefined,
      label,
      text,
      hasImageToken,
    };
  });
}


/* =========================================================
   FIND ELEMENT GEOMETRY
========================================================= */

function getElementGeometry(
  element: Element,
  svgWidth: number,
  svgHeight: number,
) {
  const tag = element.tagName.toLowerCase();

  const fallbackWidth = svgWidth * 0.28;
  const fallbackHeight = fallbackWidth / 0.72;

  /*
   * IMAGE / RECT
   */
  if (tag === 'image' || tag === 'rect') {
    const width = numeric(
      element.getAttribute('width'),
      fallbackWidth,
    );

    const height = numeric(
      element.getAttribute('height'),
      fallbackHeight,
    );

    const x = numeric(
      element.getAttribute('x'),
      (svgWidth - width) / 2,
    );

    const y = numeric(
      element.getAttribute('y'),
      (svgHeight - height) / 2,
    );

    return {
      x,
      y,
      width,
      height,
    };
  }

  /*
   * CIRCLE
   */
  if (tag === 'circle') {
    const radius = numeric(
      element.getAttribute('r'),
      Math.min(svgWidth, svgHeight) * 0.1,
    );

    const cx = numeric(
      element.getAttribute('cx'),
      svgWidth / 2,
    );

    const cy = numeric(
      element.getAttribute('cy'),
      svgHeight / 2,
    );

    return {
      x: cx - radius,
      y: cy - radius,
      width: radius * 2,
      height: radius * 2,
    };
  }

  /*
   * ELLIPSE
   */
  if (tag === 'ellipse') {
    const rx = numeric(
      element.getAttribute('rx'),
      svgWidth * 0.1,
    );

    const ry = numeric(
      element.getAttribute('ry'),
      svgHeight * 0.1,
    );

    const cx = numeric(
      element.getAttribute('cx'),
      svgWidth / 2,
    );

    const cy = numeric(
      element.getAttribute('cy'),
      svgHeight / 2,
    );

    return {
      x: cx - rx,
      y: cy - ry,
      width: rx * 2,
      height: ry * 2,
    };
  }

  /*
   * FALLBACK
   */
  return {
    x: (svgWidth - fallbackWidth) / 2,
    y: (svgHeight - fallbackHeight) / 2,
    width: fallbackWidth,
    height: fallbackHeight,
  };
}


/* =========================================================
   PARSE SVG TEMPLATE
========================================================= */

export function parseSvgTemplate(
  rawSvg: string,
): ParsedTemplate {
  const sanitizedSvg = sanitizeCanvaSvg(rawSvg);

  const found = new Set<string>();

  /*
   * Find {{name}}, {{roll}}, {{department}}, {{image}}, etc.
   */
  for (const match of sanitizedSvg.matchAll(
    /\{\{\s*([^}]+?)\s*\}\}/g,
  )) {
    found.add(cleanToken(match[1]));
  }

  const doc = new DOMParser().parseFromString(
    sanitizedSvg,
    'image/svg+xml',
  );

  const root = doc.documentElement;

  const viewBox = root
    .getAttribute('viewBox')
    ?.split(/\s+/)
    .map(Number);

  const viewWidth =
    viewBox?.[2] ??
    numeric(root.getAttribute('width'), 640);

  const viewHeight =
    viewBox?.[3] ??
    numeric(root.getAttribute('height'), 900);

  const elements = Array.from(
    doc.querySelectorAll('*'),
  );

  const elementOptions: SvgElementOption[] =
    elements.map((element, index) => {
      const tag =
        element.tagName.toLowerCase();

      const id =
        element.getAttribute('id') ?? '';

      const text =
        element.textContent?.trim() ?? '';

      const hasImageToken =
        /\{\{\s*(image|photo)\s*\}\}/i.test(text);

      let label = '';

      if (id) {
        label = `#${id}`;
      } else if (text) {
        label = text.slice(0, 60);
      } else {
        label = `<${tag}>`;
      }

      return {
        index,
        tag,
        id: id || undefined,
        label,
        text,
        hasImageToken,
      };
    });

  /*
   * IMPORTANT:
   *
   * ONLY {{image}} or {{photo}} identifies an automatic
   * placeholder.
   *
   * We DO NOT check whether the element is <image>.
   */

  let photoIndex = -1;

  for (let i = 0; i < elements.length; i++) {
    const element = elements[i];

    const text =
      element.textContent?.trim() ?? '';

    if (
      /\{\{\s*(image|photo)\s*\}\}/i.test(text)
    ) {
      photoIndex = i;
      break;
    }
  }

  let photoPlaceholder: PhotoPlaceholder | undefined;

  if (photoIndex >= 0) {
    const element = elements[photoIndex];

    const geometry = getElementGeometry(
      element,
      viewWidth,
      viewHeight,
    );

    photoPlaceholder = {
      elementIndex: photoIndex,
      ...geometry,
      source: element.tagName.toLowerCase(),
    };

    found.add('image');
  }

  return {
    rawSvg: sanitizedSvg,
    tokens: Array.from(found),
    photoPlaceholder,
    elements: elementOptions,
  };
}


/* =========================================================
   MANUALLY CREATE PHOTO PLACEHOLDER
========================================================= */

/**
 * Call this when the user manually selects an SVG element
 * from the UI.
 *
 * Example:
 *
 * const template = setPhotoPlaceholder(template, 12);
 */

export function setPhotoPlaceholder(
  template: ParsedTemplate,
  elementIndex: number,
): ParsedTemplate {
  const doc = new DOMParser().parseFromString(
    template.rawSvg,
    'image/svg+xml',
  );

  const elements = Array.from(
    doc.querySelectorAll('*'),
  );

  const element = elements[elementIndex];

  if (!element) {
    return {
      ...template,
      photoPlaceholder: undefined,
    };
  }

  const root = doc.documentElement;

  const viewBox = root
    .getAttribute('viewBox')
    ?.split(/\s+/)
    .map(Number);

  const viewWidth =
    viewBox?.[2] ??
    numeric(root.getAttribute('width'), 640);

  const viewHeight =
    viewBox?.[3] ??
    numeric(root.getAttribute('height'), 900);

  const geometry = getElementGeometry(
    element,
    viewWidth,
    viewHeight,
  );

  return {
    ...template,

    photoPlaceholder: {
      elementIndex,
      ...geometry,
      source: element.tagName.toLowerCase(),
    },
  };
}


/* =========================================================
   SPREADSHEET PARSER
========================================================= */

export async function parseSpreadsheet(
  file: File,
): Promise<SpreadsheetData> {

  /*
   * CSV
   */
  if (
    file.name
      .toLowerCase()
      .endsWith('.csv')
  ) {
    const text = await file.text();

    const parsed =
      Papa.parse<SpreadsheetRow>(
        text,
        {
          header: true,
          skipEmptyLines: true,

          transformHeader: (header) =>
            header.trim(),
        },
      );

    const rows =
      parsed.data.map((row) =>
        Object.fromEntries(
          Object.entries(row).map(
            ([key, value]) => [
              key,
              String(value ?? '').trim(),
            ],
          ),
        ),
      );

    const headers =
      parsed.meta.fields?.filter(Boolean) ??
      Object.keys(rows[0] ?? {});

    return {
      headers,
      rows,
    };
  }

  /*
   * XLSX
   */
  const buffer =
    await file.arrayBuffer();

  const workbook =
    XLSX.read(buffer, {
      type: 'array',
    });

  const firstSheet =
    workbook.Sheets[
      workbook.SheetNames[0]
    ];

  const rows =
    XLSX.utils
      .sheet_to_json<SpreadsheetRow>(
        firstSheet,
        {
          defval: '',
          raw: false,
        },
      )
      .map((row) =>
        Object.fromEntries(
          Object.entries(row).map(
            ([key, value]) => [
              key.trim(),
              String(value ?? '').trim(),
            ],
          ),
        ),
      );

  const embeddedPhotos:
    Record<number, File> = {};

  /*
   * Attempt to read Excel embedded photos.
   */
  try {
    const zip =
      await JSZip.loadAsync(buffer);

    const richXml =
      await zip
        .file(
          'xl/richData/rdrichvalue.xml',
        )
        ?.async('text');

    const sheetXml =
      await zip
        .file(
          'xl/worksheets/sheet1.xml',
        )
        ?.async('text');

    if (richXml && sheetXml) {
      const richDoc =
        new DOMParser().parseFromString(
          richXml,
          'application/xml',
        );

      const richValues =
        Array.from(
          richDoc.getElementsByTagNameNS(
            '*',
            'rv',
          ),
        );

      const sheetDoc =
        new DOMParser().parseFromString(
          sheetXml,
          'application/xml',
        );

      const cells =
        Array.from(
          sheetDoc.getElementsByTagNameNS(
            '*',
            'c',
          ),
        );

      const mediaFiles =
        new Set(
          Object.keys(zip.files).filter(
            (name) =>
              /^xl\/media\/image\d+\.(png|jpe?g|webp|gif)$/i.test(
                name,
              ),
          ),
        );

      for (const cell of cells) {
        const vm =
          cell.getAttribute('vm');

        const ref =
          cell.getAttribute('r');

        if (
          !vm ||
          !ref ||
          !/^[A-Z]+\d+$/.test(ref)
        ) {
          continue;
        }

        const richIndex =
          Number(vm) - 1;

        if (
          !Number.isInteger(richIndex) ||
          richIndex < 0 ||
          richIndex >= richValues.length
        ) {
          continue;
        }

        const rich =
          richValues[richIndex];

        const values =
          Array.from(
            rich.getElementsByTagNameNS(
              '*',
              'v',
            ),
          ).map(
            (node) =>
              node.textContent ?? '',
          );

        const localImageId =
          Number(values[0]);

        if (
          !Number.isInteger(
            localImageId,
          )
        ) {
          continue;
        }

        const mediaCandidates =
          Array.from(mediaFiles).filter(
            (name) =>
              Number(
                name.match(
                  /image(\d+)\./i,
                )?.[1],
              ) ===
              localImageId + 1,
          );

        const mediaName =
          mediaCandidates[0];

        if (!mediaName) {
          continue;
        }

        const blob =
          await zip
            .file(mediaName)!
            .async('blob');

        const extension =
          mediaName
            .split('.')
            .pop()
            ?.toLowerCase() ??
          'png';

        const mime =
          extension === 'jpg' ||
          extension === 'jpeg'
            ? 'image/jpeg'
            : extension === 'webp'
              ? 'image/webp'
              : extension === 'gif'
                ? 'image/gif'
                : 'image/png';

        const rowNumber =
          Number(
            ref.match(/\d+$/)?.[0],
          );

        const dataRowIndex =
          rowNumber - 2;

        if (
          dataRowIndex >= 0 &&
          dataRowIndex < rows.length
        ) {
          embeddedPhotos[
            dataRowIndex
          ] = new File(
            [blob],
            mediaName
              .split('/')
              .pop() ?? mediaName,
            {
              type: mime,
            },
          );
        }
      }
    }
  } catch {
    /*
     * Embedded image extraction failing
     * should NOT break spreadsheet parsing.
     */
  }

  return {
    headers: Object.keys(
      rows[0] ?? {},
    ),
    rows,
    embeddedPhotos,
  };
}


/* =========================================================
   PHOTO ARCHIVE
========================================================= */

export async function parsePhotoArchive(
  file: File,
): Promise<PhotoArchive> {

  const zip =
    await JSZip.loadAsync(file);

  const files:
    Record<string, File> = {};

  const names: string[] = [];

  await Promise.all(
    Object.values(zip.files)
      .filter(
        (entry) => !entry.dir,
      )
      .map(async (entry) => {

        const name =
          entry.name;

        if (
          !/\.(png|jpe?g|webp|gif)$/i.test(
            name,
          )
        ) {
          return;
        }

        const blob =
          await entry.async('blob');

        const extension =
          name
            .split('.')
            .pop()
            ?.toLowerCase();

        const inferredType =
          extension === 'jpg' ||
          extension === 'jpeg'
            ? 'image/jpeg'
            : extension === 'webp'
              ? 'image/webp'
              : extension === 'gif'
                ? 'image/gif'
                : 'image/png';

        const imageFile =
          new File(
            [blob],
            name.split('/').pop() ??
              name,
            {
              type:
                blob.type ||
                inferredType,
            },
          );

        files[
          fileNameKey(name)
        ] = imageFile;

        files[
          baseName(name)
        ] = imageFile;

        names.push(
          name.split('/').pop() ??
            name,
        );
      }),
  );

  return {
    files,
    names: names.sort(),
  };
}


/* =========================================================
   FIND PHOTO
========================================================= */

export function findPhoto(
  value: string,
  archive?: PhotoArchive,
): File | string | undefined {

  const trimmed =
    value.trim();

  if (!trimmed) {
    return undefined;
  }

  if (
    /^https?:\/\//i.test(trimmed) ||
    trimmed.startsWith('data:')
  ) {
    return trimmed;
  }

  if (!archive) {
    return undefined;
  }

  return (
    archive.files[
      fileNameKey(trimmed)
    ] ??
    archive.files[
      baseName(trimmed)
    ]
  );
}


/* =========================================================
   FILE → DATA URL
========================================================= */

export async function fileToDataUrl(
  file: File,
): Promise<string> {

  return await new Promise(
    (resolve, reject) => {

      const reader =
        new FileReader();

      reader.onload = () =>
        resolve(
          String(
            reader.result,
          ),
        );

      reader.onerror = () =>
        reject(
          reader.error ??
            new Error(
              'Could not read image',
            ),
        );

      reader.readAsDataURL(file);
    },
  );
}


/* =========================================================
   ROW LABEL
========================================================= */

export function rowLabel(
  row: SpreadsheetRow,
  mappings: FieldMapping,
  index: number,
) {

  const identityToken =
    Object.keys(mappings).find(
      (token) =>
        /name|person|student|roll|id|identifier/i.test(
          token,
        ) &&
        mappings[token] &&
        mappings[token] !==
          '__ignore__',
    );

  const preferred =
    identityToken
      ? mappings[identityToken]
      : undefined;

  return (
    (preferred &&
      row[preferred]) ||
    row.Name ||
    row.name ||
    row['Full name'] ||
    row.ID ||
    row.Id ||
    `Row ${index + 1}`
  );
}


/* =========================================================
   REPLACE PHOTO ELEMENT
========================================================= */

/**
 * Converts the selected SVG element into a real <image>
 * if necessary.
 *
 * The exact geometry of the selected element is preserved.
 */

function replaceElementWithImage(
  doc: Document,
  selectedElement: Element,
  photoData: string,
  placeholder: PhotoPlaceholder,
) {

  const tag =
    selectedElement.tagName.toLowerCase();

  let imageElement: Element;

  /*
   * If user selected an existing <image>,
   * simply use it.
   */
  if (tag === 'image') {
    imageElement =
      selectedElement;
  } else {

    /*
     * Otherwise convert the selected shape
     * into an actual SVG <image>.
     */

    imageElement =
      doc.createElementNS(
        'http://www.w3.org/2000/svg',
        'image',
      );

    /*
     * Preserve exact geometry.
     */

    const geometry =
      getElementGeometry(
        selectedElement,
        placeholder.x +
          placeholder.width,
        placeholder.y +
          placeholder.height,
      );

    imageElement.setAttribute(
      'x',
      String(
        selectedElement.getAttribute(
          'x',
        ) ??
          geometry.x ??
          placeholder.x,
      ),
    );

    imageElement.setAttribute(
      'y',
      String(
        selectedElement.getAttribute(
          'y',
        ) ??
          geometry.y ??
          placeholder.y,
      ),
    );

    imageElement.setAttribute(
      'width',
      String(
        selectedElement.getAttribute(
          'width',
        ) ??
          geometry.width ??
          placeholder.width,
      ),
    );

    imageElement.setAttribute(
      'height',
      String(
        selectedElement.getAttribute(
          'height',
        ) ??
          geometry.height ??
          placeholder.height,
      ),
    );

    /*
     * Preserve important attributes.
     */

    for (
      const attribute of [
        'id',
        'class',
        'clip-path',
        'mask',
        'transform',
        'opacity',
      ]
    ) {
      const value =
        selectedElement.getAttribute(
          attribute,
        );

      if (value) {
        imageElement.setAttribute(
          attribute,
          value,
        );
      }
    }

    /*
     * If it was a circle/ellipse and had no
     * clipping information, create a clip
     * so the image stays inside the shape.
     */

    if (
      (tag === 'circle' ||
        tag === 'ellipse') &&
      !selectedElement.getAttribute(
        'clip-path',
      )
    ) {

      const clipId =
        `photo-clip-${placeholder.elementIndex}`;

      const defs =
        doc.querySelector('defs') ??
        doc.documentElement.insertBefore(
          doc.createElementNS(
            'http://www.w3.org/2000/svg',
            'defs',
          ),
          doc.documentElement.firstChild,
        );

      const clipPath =
        doc.createElementNS(
          'http://www.w3.org/2000/svg',
          'clipPath',
        );

      clipPath.setAttribute(
        'id',
        clipId,
      );

      const shape =
        selectedElement.cloneNode(
          true,
        ) as Element;

      shape.removeAttribute('id');

      clipPath.appendChild(shape);

      defs.appendChild(
        clipPath,
      );

      imageElement.setAttribute(
        'clip-path',
        `url(#${clipId})`,
      );
    }

    /*
     * Replace ONLY this element.
     */

    selectedElement.parentNode?.replaceChild(
      imageElement,
      selectedElement,
    );
  }

  /*
   * Put the actual photo here.
   */

  imageElement.setAttribute(
    'href',
    photoData,
  );

  imageElement.setAttribute(
    'preserveAspectRatio',
    'xMidYMid slice',
  );

  imageElement.setAttribute(
    'data-generated-photo',
    'true',
  );

  /*
   * Remove old xlink reference if present.
   */

  imageElement.removeAttribute(
    'xlink:href',
  );

  return imageElement;
}


/* =========================================================
   RENDER SVG FOR ROW
========================================================= */

function renderSvgForRowBase(
  template: ParsedTemplate,
  row: SpreadsheetRow,
  mappings: FieldMapping,
  photoData?: string,

  /*
   * Optional manual selection.
   *
   * If provided, this ALWAYS wins over {{image}}.
   */
  selectedPhotoElementIndex?: number,
) {

  /*
   * Replace normal {{tokens}}.
   *
   * {{image}} and {{photo}} are removed because
   * they are placeholders, not text.
   */

  let svg =
    template.rawSvg.replace(
      /\{\{\s*([^}]+?)\s*\}\}/g,
      (
        _match,
        rawToken: string,
      ) => {

        const token =
          cleanToken(rawToken);

        /*
         * PHOTO TOKENS
         */

        if (
          token.toLowerCase() ===
            'image' ||
          token.toLowerCase() ===
            'photo'
        ) {
          return '';
        }

        /*
         * NORMAL FIELD
         */

        const header =
          mappings[token];

        return header &&
          header !== '__ignore__'
          ? escapeXml(
              row[header] ?? '',
            )
          : '';
      },
    );

  /*
   * No photo supplied → just return SVG.
   */

  if (!photoData) {
    return svg;
  }

  /*
   * Parse generated SVG.
   */

  const doc =
    new DOMParser().parseFromString(
      svg,
      'image/svg+xml',
    );

  const elements =
    Array.from(
      doc.querySelectorAll('*'),
    );

  /*
   * -------------------------------------------------------
   * STEP 1
   *
   * If user manually selected an element,
   * use THAT element.
   * -------------------------------------------------------
   */

  let photoIndex =
    selectedPhotoElementIndex;

  /*
   * -------------------------------------------------------
   * STEP 2
   *
   * If no manual selection was made,
   * search ONLY for {{image}} / {{photo}}.
   * -------------------------------------------------------
   */

  if (
    photoIndex == null
  ) {

    photoIndex =
      elements.findIndex(
        (element) =>
          /\{\{\s*(image|photo)\s*\}\}/i.test(
            element.textContent ??
              '',
          ),
      );
  }

  /*
   * No placeholder found.
   *
   * IMPORTANT:
   * We DO NOT fall back to the first <image>.
   */

  if (
    photoIndex == null ||
    photoIndex < 0 ||
    photoIndex >= elements.length
  ) {
    return svg;
  }

  const photoElement =
    elements[photoIndex];

  /*
   * Get exact geometry before replacing.
   */

  const root =
    doc.documentElement;

  const viewBox =
    root
      .getAttribute('viewBox')
      ?.split(/\s+/)
      .map(Number);

  const viewWidth =
    viewBox?.[2] ??
    numeric(
      root.getAttribute(
        'width',
      ),
      640,
    );

  const viewHeight =
    viewBox?.[3] ??
    numeric(
      root.getAttribute(
        'height',
      ),
      900,
    );

  const geometry =
    getElementGeometry(
      photoElement,
      viewWidth,
      viewHeight,
    );

  const placeholder:
    PhotoPlaceholder = {
      elementIndex: photoIndex,
      ...geometry,
      source:
        photoElement.tagName.toLowerCase(),
    };

  /*
   * Replace ONLY selected element.
   */

  replaceElementWithImage(
    doc,
    photoElement,
    photoData,
    placeholder,
  );

  /*
   * Serialize back to SVG.
   */

  svg =
    new XMLSerializer()
      .serializeToString(doc);

  return svg;
}


/* =========================================================
   TEXT FITTING
   Long values (e.g. a long Department name) are wrapped onto a
   second line when the template leaves room below, otherwise the
   font is shrunk. The photo placeholder is treated as an obstacle,
   so text never runs underneath the photo.
========================================================= */

let measureFontsReady = false;
let measureContext: CanvasRenderingContext2D | null = null;

function preloadMeasureFonts() {
  if (typeof document === 'undefined' || !('fonts' in document)) return;
  const base = (import.meta as any).env?.BASE_URL ?? '/';
  Promise.all(
    FONT_FACES.map(async (face) => {
      const fontFace = new FontFace(
        face.family,
        `url(${base}fonts/${face.file})`,
        { weight: String(face.weight), style: 'normal' },
      );
      await fontFace.load();
      document.fonts.add(fontFace);
    }),
  )
    .then(() => {
      measureFontsReady = true;
    })
    .catch(() => {
      /* fall back to estimated widths */
    });
}

function measureTextWidth(
  text: string,
  family: string,
  weight: number,
  fontSize: number,
  spacingEm: number,
): number {
  const spacing = spacingEm * fontSize * text.length;

  if (measureFontsReady && typeof document !== 'undefined') {
    if (!measureContext) {
      measureContext = document.createElement('canvas').getContext('2d');
    }
    if (measureContext) {
      measureContext.font = `${weight} ${fontSize}px "${family}"`;
      return measureContext.measureText(text).width + spacing;
    }
  }

  return text.length * fontSize * 0.56 + spacing;
}

function inheritedAttr(el: Element, name: string): string | null {
  let node: Element | null = el;
  while (node) {
    const direct = node.getAttribute(name);
    if (direct) return direct;
    const style = node.getAttribute('style');
    if (style) {
      const match = style.match(
        new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*([^;]+)`, 'i'),
      );
      if (match) return match[1].trim();
    }
    node = node.parentElement;
  }
  return null;
}

function hasTransform(el: Element): boolean {
  let node: Element | null = el;
  while (node) {
    if (node.getAttribute('transform')) return true;
    node = node.parentElement;
  }
  return false;
}

function parseFontWeight(value: string | null): number {
  if (!value) return 400;
  const v = value.trim().toLowerCase();
  if (v === 'bold' || v === 'bolder') return 700;
  if (v === 'normal') return 400;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n : 400;
}

type FitObstacle = {
  shape: 'ellipse' | 'rect';
  x: number;
  y: number;
  width: number;
  height: number;
};

type FitItem = {
  index: number;
  raw: string;
  x: number;
  y: number;
  fontSize: number;
  family: string;
  weight: number;
  spacingEm: number;
  canWrap: boolean;
};

type FitAnalysis = {
  viewWidth: number;
  obstacle: FitObstacle | null;
  items: FitItem[];
};

const TEXT_FIT_LINE_HEIGHT = 1.3;
const TEXT_FIT_GAP = 24;
const fitCache = new WeakMap<object, Map<string, FitAnalysis>>();

function obstacleLeft(
  obstacle: FitObstacle,
  top: number,
  bottom: number,
): number | null {
  if (bottom < obstacle.y || top > obstacle.y + obstacle.height) return null;
  if (obstacle.shape === 'rect') return obstacle.x;

  const rx = obstacle.width / 2;
  const ry = obstacle.height / 2;
  const cx = obstacle.x + rx;
  const cy = obstacle.y + ry;
  const dy =
    cy >= top && cy <= bottom
      ? 0
      : Math.min(Math.abs(top - cy), Math.abs(bottom - cy));
  if (dy >= ry) return null;
  return cx - rx * Math.sqrt(1 - (dy / ry) ** 2);
}

function availableWidth(
  analysis: FitAnalysis,
  item: FitItem,
  lines: number,
  fontSize: number,
): number {
  const lineHeight = fontSize * TEXT_FIT_LINE_HEIGHT;
  const top = item.y - fontSize * 0.8;
  const bottom = item.y + (lines - 1) * lineHeight + fontSize * 0.25;
  let right = analysis.viewWidth * 0.9;

  if (analysis.obstacle) {
    const left = obstacleLeft(analysis.obstacle, top, bottom);
    if (left != null && left > item.x) {
      right = Math.min(right, left - TEXT_FIT_GAP);
    }
  }
  return right - item.x;
}

function getFitAnalysis(
  template: ParsedTemplate,
  selectedPhotoElementIndex?: number,
): FitAnalysis | null {
  const cacheKey = String(selectedPhotoElementIndex ?? 'auto');
  const cached = fitCache.get(template)?.get(cacheKey);
  if (cached) return cached;

  const doc = new DOMParser().parseFromString(
    template.rawSvg,
    'image/svg+xml',
  );
  if (doc.querySelector('parsererror')) return null;

  const root = doc.documentElement;
  const viewBox = root
    .getAttribute('viewBox')
    ?.split(/\s+/)
    .map(Number);
  const viewWidth = viewBox?.[2] ?? numeric(root.getAttribute('width'), 640);
  const viewHeight = viewBox?.[3] ?? numeric(root.getAttribute('height'), 900);

  /* photo placeholder geometry (obstacle) */
  const elements = Array.from(doc.querySelectorAll('*'));
  let photoIndex = selectedPhotoElementIndex;
  if (photoIndex == null) {
    photoIndex = elements.findIndex((el) =>
      /\{\{\s*(image|photo)\s*\}\}/i.test(el.textContent ?? ''),
    );
  }
  let obstacle: FitObstacle | null = null;
  const photoEl =
    photoIndex != null && photoIndex >= 0 ? elements[photoIndex] : undefined;
  if (photoEl) {
    const tag = photoEl.tagName.toLowerCase();
    if (['ellipse', 'circle', 'rect', 'image'].includes(tag)) {
      obstacle = {
        shape: tag === 'ellipse' || tag === 'circle' ? 'ellipse' : 'rect',
        ...getElementGeometry(photoEl, viewWidth, viewHeight),
      };
    }
  }

  /* every <text> element, with its position and font */
  type Basic = Omit<FitItem, 'canWrap'> & { hasToken: boolean; width: number };
  const basics: Basic[] = [];
  Array.from(doc.querySelectorAll('text')).forEach((textEl, index) => {
    if (hasTransform(textEl)) return;
    const anchor = inheritedAttr(textEl, 'text-anchor');
    if (anchor === 'middle' || anchor === 'end') return;

    const first = textEl.querySelector('tspan');
    const x = Number.parseFloat(
      first?.getAttribute('x') ?? textEl.getAttribute('x') ?? '',
    );
    const y = Number.parseFloat(
      first?.getAttribute('y') ?? textEl.getAttribute('y') ?? '',
    );
    if (!Number.isFinite(x) || !Number.isFinite(y)) return;

    const raw = (textEl.textContent ?? '').trim();
    const fontSize = numeric(inheritedAttr(textEl, 'font-size'), 16);
    const family = (inheritedAttr(textEl, 'font-family') ?? 'sans-serif')
      .split(',')[0]
      .replace(/["']/g, '')
      .trim();
    const weight = parseFontWeight(inheritedAttr(textEl, 'font-weight'));
    const spacingRaw = (inheritedAttr(textEl, 'letter-spacing') ?? '').trim();
    const spacingEm = spacingRaw.endsWith('em')
      ? numeric(spacingRaw, 0)
      : spacingRaw.endsWith('px')
        ? numeric(spacingRaw, 0) / fontSize
        : 0;

    const hasToken = Array.from(
      raw.matchAll(/\{\{\s*([^}]+?)\s*\}\}/g),
    ).some((m) => !['image', 'photo'].includes(m[1].trim().toLowerCase()));

    basics.push({
      index,
      raw,
      x,
      y,
      fontSize,
      family,
      weight,
      spacingEm,
      hasToken,
      width: measureTextWidth(raw, family, weight, fontSize, spacingEm),
    });
  });

  const analysis: FitAnalysis = {
    viewWidth,
    obstacle,
    items: basics
      .filter((b) => b.hasToken)
      .map((b) => {
        const lineHeight = b.fontSize * TEXT_FIT_LINE_HEIGHT;
        const rightEdge = b.x + Math.max(availableWidth(
          { viewWidth, obstacle, items: [] },
          { ...b, canWrap: false },
          1,
          b.fontSize,
        ), 0);

        /* nearest text below that overlaps horizontally */
        let nextTop = Number.POSITIVE_INFINITY;
        for (const other of basics) {
          if (other.index === b.index || other.y <= b.y + 1) continue;
          if (other.x > rightEdge || other.x + other.width < b.x) continue;
          nextTop = Math.min(nextTop, other.y - other.fontSize * 0.75);
        }
        const canWrap =
          b.y + lineHeight + b.fontSize * 0.2 <= nextTop - 4;

        return {
          index: b.index,
          raw: b.raw,
          x: b.x,
          y: b.y,
          fontSize: b.fontSize,
          family: b.family,
          weight: b.weight,
          spacingEm: b.spacingEm,
          canWrap,
        };
      }),
  };

  /* do not cache estimates made before fonts finished loading */
  if (measureFontsReady) {
    const perTemplate = fitCache.get(template) ?? new Map<string, FitAnalysis>();
    perTemplate.set(cacheKey, analysis);
    fitCache.set(template, perTemplate);
  }
  return analysis;
}

function resolveTokenText(
  raw: string,
  row: SpreadsheetRow,
  mappings: FieldMapping,
): string {
  return raw.replace(
    /\{\{\s*([^}]+?)\s*\}\}/g,
    (_match, rawToken: string) => {
      const token = cleanToken(rawToken);
      if (['image', 'photo'].includes(token.toLowerCase())) return '';
      const header = mappings[token];
      return header && header !== '__ignore__' ? String(row[header] ?? '') : '';
    },
  );
}

function planTextFit(
  analysis: FitAnalysis,
  item: FitItem,
  value: string,
): { lines: string[]; scale: number } | null {
  const text = value.replace(/\s+/g, ' ').trim();
  if (!text) return null;

  const width = (s: string, fs: number) =>
    measureTextWidth(s, item.family, item.weight, fs, item.spacingEm);

  if (width(text, item.fontSize) <= availableWidth(analysis, item, 1, item.fontSize)) {
    return null;
  }

  const words = text.split(' ');
  for (let step = 0; step <= 10; step += 1) {
    const scale = 1 - step * 0.04;
    const fs = item.fontSize * scale;

    if (width(text, fs) <= availableWidth(analysis, item, 1, fs)) {
      return { lines: [text], scale };
    }

    if (item.canWrap && words.length > 1) {
      const limit = availableWidth(analysis, item, 2, fs);
      let best: string[] | null = null;
      let bestMax = Number.POSITIVE_INFINITY;
      for (let k = 1; k < words.length; k += 1) {
        const a = words.slice(0, k).join(' ');
        const b = words.slice(k).join(' ');
        const longest = Math.max(width(a, fs), width(b, fs));
        if (longest < bestMax) {
          bestMax = longest;
          best = [a, b];
        }
      }
      if (best && bestMax <= limit) return { lines: best, scale };
    }
  }

  return { lines: [text], scale: 0.6 };
}

function fitLongText(
  svg: string,
  template: ParsedTemplate,
  row: SpreadsheetRow,
  mappings: FieldMapping,
  selectedPhotoElementIndex?: number,
): string {
  const analysis = getFitAnalysis(template, selectedPhotoElementIndex);
  if (!analysis || !analysis.items.length) return svg;

  const plans: { item: FitItem; lines: string[]; scale: number }[] = [];
  for (const item of analysis.items) {
    const value = resolveTokenText(item.raw, row, mappings);
    const plan = planTextFit(analysis, item, value);
    if (plan) plans.push({ item, ...plan });
  }
  if (!plans.length) return svg;

  const doc = new DOMParser().parseFromString(svg, 'image/svg+xml');
  if (doc.querySelector('parsererror')) return svg;
  const texts = doc.querySelectorAll('text');
  const ns = 'http://www.w3.org/2000/svg';

  for (const { item, lines, scale } of plans) {
    const el = texts[item.index];
    if (!el) continue;

    const fontSize = item.fontSize * scale;
    const lineHeight = fontSize * TEXT_FIT_LINE_HEIGHT;

    if (scale < 1) {
      const style = el.getAttribute('style');
      if (style && /font-size\s*:/i.test(style)) {
        el.setAttribute(
          'style',
          style.replace(/font-size\s*:[^;]+/i, `font-size: ${fontSize}px`),
        );
      } else {
        el.setAttribute('font-size', String(fontSize));
      }
    }

    while (el.firstChild) el.removeChild(el.firstChild);
    lines.forEach((line, i) => {
      const tspan = doc.createElementNS(ns, 'tspan');
      tspan.setAttribute('x', String(item.x));
      tspan.setAttribute('y', String(item.y + i * lineHeight));
      tspan.textContent = line;
      el.appendChild(tspan);
    });
  }

  return new XMLSerializer().serializeToString(doc);
}

export function renderSvgForRow(
  template: ParsedTemplate,
  row: SpreadsheetRow,
  mappings: FieldMapping,
  photoData?: string,
  selectedPhotoElementIndex?: number,
) {
  const svg = renderSvgForRowBase(
    template,
    row,
    mappings,
    photoData,
    selectedPhotoElementIndex,
  );

  try {
    return fitLongText(svg, template, row, mappings, selectedPhotoElementIndex);
  } catch (error) {
    console.warn('Text fitting skipped', error);
    return svg;
  }
}


/* =========================================================
   SVG DIMENSIONS
========================================================= */

function svgDimensions(
  svgText: string,
) {

  const doc =
    new DOMParser().parseFromString(
      svgText,
      'image/svg+xml',
    );

  const root =
    doc.documentElement;

  const viewBox =
    root
      .getAttribute('viewBox')
      ?.split(/\s+/)
      .map(Number);

  return {
    width:
      viewBox?.[2] ||
      numeric(
        root.getAttribute(
          'width',
        ),
        640,
      ),

    height:
      viewBox?.[3] ||
      numeric(
        root.getAttribute(
          'height',
        ),
        900,
      ),
  };
}


/* =========================================================
   FONT EMBEDDING (for SVG -> PNG)
   An SVG loaded through <img>/Image runs in an isolated context
   and cannot see page fonts, so fonts must be embedded as data URLs.
   Font files live in public/fonts.
========================================================= */

const FONT_FACES = [
  { family: 'Lato', weight: 700, file: 'lato-latin-700-normal.woff2' },
  { family: 'Public Sans', weight: 400, file: 'public-sans-latin-400-normal.woff2' },
  { family: 'Public Sans', weight: 700, file: 'public-sans-latin-700-normal.woff2' },
  { family: 'Source Serif 4', weight: 700, file: 'source-serif-4-latin-700-normal.woff2' },
];

/* Preload the same fonts into the page so text can be measured for wrapping. */
preloadMeasureFonts();

let fontCssPromise: Promise<string> | null = null;

async function fontFileToDataUrl(file: string): Promise<string> {
  const base = (import.meta as any).env?.BASE_URL ?? '/';
  const response = await fetch(`${base}fonts/${file}`);
  if (!response.ok) throw new Error(`Font ${file} could not be loaded`);
  const blob = await response.blob();
  return await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

function loadFontCss(): Promise<string> {
  if (!fontCssPromise) {
    fontCssPromise = Promise.all(
      FONT_FACES.map(async (face) => {
        const dataUrl = await fontFileToDataUrl(face.file);
        return `@font-face{font-family:'${face.family}';font-weight:${face.weight};font-style:normal;src:url(${dataUrl}) format('woff2');}`;
      }),
    )
      .then((rules) => rules.join(''))
      .catch((error) => {
        // Do not cache a failure, and do not block export.
        fontCssPromise = null;
        console.warn('Font embedding failed, exporting with fallback fonts', error);
        return '';
      });
  }
  return fontCssPromise;
}

async function embedFonts(svgText: string): Promise<string> {
  const css = await loadFontCss();
  if (!css) return svgText;
  return svgText.replace(
    /(<svg[^>]*>)/,
    `$1<defs><style type="text/css">${css}</style></defs>`,
  );
}


/* =========================================================
   SVG → PNG
========================================================= */

export async function renderSvgToPng(
  svgText: string,
  scale = 2,
): Promise<Blob> {

  const {
    width,
    height,
  } =
    svgDimensions(svgText);

  const canvas =
    document.createElement(
      'canvas',
    );

  canvas.width =
    Math.max(
      1,
      Math.round(
        width * scale,
      ),
    );

  canvas.height =
    Math.max(
      1,
      Math.round(
        height * scale,
      ),
    );

  const context =
    canvas.getContext('2d');

  if (!context) {
    throw new Error(
      'Canvas rendering is not supported in this browser',
    );
  }

  context.fillStyle =
    '#fffdf8';

  context.fillRect(
    0,
    0,
    canvas.width,
    canvas.height,
  );

  const image =
    new Image();

  const svgWithFonts =
    await embedFonts(svgText);

  const svgUrl =
    `data:image/svg+xml;charset=utf-8,${encodeURIComponent(
      svgWithFonts,
    )}`;

  await new Promise<void>(
    (
      resolve,
      reject,
    ) => {

      image.onload =
        () => resolve();

      image.onerror =
        () =>
          reject(
            new Error(
              'The SVG preview could not be rendered',
            ),
          );

      image.src = svgUrl;
    },
  );

  context.drawImage(
    image,
    0,
    0,
    canvas.width,
    canvas.height,
  );

  return await new Promise<Blob>(
    (
      resolve,
      reject,
    ) => {

      canvas.toBlob(
        (blob) => {

          if (blob) {
            resolve(blob);
          } else {
            reject(
              new Error(
                'PNG export failed',
              ),
            );
          }
        },
        'image/png',
        0.96,
      );
    },
  );
}


/* =========================================================
   RESOLVE PHOTO DATA
========================================================= */

export async function resolvePhotoData(
  value: string,
  archive?: PhotoArchive,
): Promise<
  string | undefined
> {

  const result =
    findPhoto(
      value,
      archive,
    );

  if (!result) {
    return undefined;
  }

  /*
   * Already a data URL.
   */

  if (
    typeof result === 'string'
  ) {

    if (
      result.startsWith(
        'data:',
      )
    ) {
      return result;
    }

    /*
     * Remote URL.
     */

    try {

      const response =
        await fetch(
          result,
          {
            mode: 'cors',
          },
        );

      if (!response.ok) {
        return undefined;
      }

      const blob =
        await response.blob();

      return await fileToDataUrl(
        new File(
          [blob],
          'remote-photo',
        ),
      );

    } catch {
      return undefined;
    }
  }

  /*
   * Local file.
   */

  return await fileToDataUrl(
    result,
  );
}
