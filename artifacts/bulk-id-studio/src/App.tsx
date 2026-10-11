import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type DragEvent,
  type ReactNode,
  type MouseEvent,
} from 'react';
import {
  AlertTriangle,
  ArrowLeft,
  ArrowRight,
  Check,
  CircleHelp,
  CloudUpload,
  Download,
  FileArchive,
  FileSpreadsheet,
  FileText,
  Image as ImageIcon,
  Layers3,
  LoaderCircle,
  MoreHorizontal,
  Play,
  Printer,
  RefreshCw,
  Search,
  Settings2,
  ShieldCheck,
  Sparkles,
  Upload,
  X,
  Zap,
} from 'lucide-react';
import JSZip from 'jszip';
import { PDFDocument } from 'pdf-lib';
import { Route, Switch } from 'wouter';

import {
  type FieldMapping,
  type GeneratedCard,
  type GenerationIssue,
  type ParsedTemplate,
  type PhotoArchive,
  type SpreadsheetData,
  type SpreadsheetRow,
  fileToDataUrl,
  parsePhotoArchive,
  parseSpreadsheet,
  parseSvgTemplate,
  renderSvgForRow,
  renderSvgToPng,
  resolvePhotoData,
  rowLabel,
  setPhotoPlaceholder,
} from '@/lib/studio';

import '@/index.css';

type Step = 'Upload' | 'Map fields' | 'Review' | 'Export';

const steps: Step[] = ['Upload', 'Map fields', 'Review', 'Export'];

const normalize = (value: string) =>
  value.toLowerCase().replace(/[^a-z0-9]/g, '');

function escapeSvgText(value: string) {
  return value.replace(
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
}

function makeSelectableSvg(svgText: string, selectedIndex: number | null) {
  const doc = new DOMParser().parseFromString(svgText, 'image/svg+xml');
  const elements = Array.from(doc.querySelectorAll('*'));

  const selectableTags = new Set([
    'rect', 'circle', 'ellipse', 'image', 'path',
    'polygon', 'polyline', 'line', 'text', 'tspan',
  ]);

  elements.forEach((element, index) => {
    const tag = element.tagName.toLowerCase();
    if (!selectableTags.has(tag)) {
      element.removeAttribute('data-bulk-id-element');
      element.removeAttribute('data-photo-selected');
      return;
    }
    element.setAttribute('data-bulk-id-element', String(index));
    if (selectedIndex === index) {
      element.setAttribute('data-photo-selected', 'true');
    } else {
      element.removeAttribute('data-photo-selected');
    }
  });

  const style = doc.createElementNS('http://www.w3.org/2000/svg', 'style');
  style.textContent = `
    [data-bulk-id-element] { cursor: crosshair; }
    [data-bulk-id-element]:hover {
      filter: drop-shadow(0 0 2px rgba(255,229,102,.9)) drop-shadow(0 0 5px rgba(255,229,102,.55));
    }
    [data-photo-selected="true"] {
      filter: drop-shadow(0 0 2px rgba(255,229,102,1)) drop-shadow(0 0 5px rgba(255,229,102,.9)) drop-shadow(0 0 9px rgba(255,229,102,.45));
    }
  `;
  doc.documentElement.insertBefore(style, doc.documentElement.firstChild);
  return new XMLSerializer().serializeToString(doc);
}

function suggestMappings(template: ParsedTemplate, spreadsheet: SpreadsheetData | null): FieldMapping {
  if (!spreadsheet) return {};
  return Object.fromEntries(
    template.tokens.map((token) => {
      const tokenKey = normalize(token);
      const header = spreadsheet.headers.find((candidate) => {
        const key = normalize(candidate);
        if (key === tokenKey) return true;
        if (tokenKey === 'photo') {
          return ['photo', 'photourl', 'image', 'imageurl', 'filename', 'photofilename'].includes(key);
        }
        return key.includes(tokenKey) || tokenKey.includes(key);
      });
      return [token, header ?? ''];
    }),
  );
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/* ========== PIXEL UI COMPONENTS ========== */

function Pill({
  children,
  tone = 'neutral',
}: {
  children: ReactNode;
  tone?: 'neutral' | 'green' | 'amber' | 'red';
}) {
  const colors = {
    neutral: 'bg-[#e8e0d0] text-[#1a1a2e]',
    green: 'bg-[#4ade80] text-[#1a1a2e]',
    amber: 'bg-[#ffe566] text-[#1a1a2e]',
    red: 'bg-[#ff6b6b] text-white',
  };
  return (
    <span className={`pixel-font inline-flex items-center pixel-border-sm px-2 py-1 text-[9px] uppercase tracking-wider ${colors[tone]}`}>
      {children}
    </span>
  );
}

function UploadZone({
  kind,
  accept,
  icon: Icon,
  title,
  detail,
  onFile,
  fileName,
}: {
  kind: string;
  accept: string;
  icon: typeof Upload;
  title: string;
  detail: string;
  onFile: (file: File) => void;
  fileName?: string;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);

  const handleDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragging(false);
    const file = event.dataTransfer.files[0];
    if (file) onFile(file);
  };

  return (
    <div
      className={`group pixel-border bg-[#fff8e7] p-3 transition-all ${
        dragging ? 'bg-[#ffe566] -translate-x-px -translate-y-px' : ''
      } ${fileName ? 'ring-2 ring-[#4ade80]' : ''}`}
      onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={handleDrop}
      data-testid={`upload-zone-${kind}`}
    >
      <input
        ref={inputRef}
        className="hidden"
        type="file"
        accept={accept}
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) onFile(file);
        }}
        data-testid={`input-file-${kind}`}
      />
      <button type="button" className="w-full text-left" onClick={() => inputRef.current?.click()} data-testid={`button-choose-${kind}`}>
        <div className="flex items-center gap-3">
          <span className="grid h-10 w-10 shrink-0 place-items-center pixel-border-sm bg-[#7ec8f5] text-[#1a1a2e]">
            <Icon size={18} strokeWidth={2.5} />
          </span>
          <span className="min-w-0">
            <span className="block text-[12px] font-bold text-[#1a1a2e]">{fileName ?? title}</span>
            <span className="mt-1 block truncate text-[10px] text-[#1a1a2e]/60">
              {fileName ? 'Replace file · click to choose another' : detail}
            </span>
          </span>
          <Upload className="ml-auto text-[#1a1a2e]/50 transition-transform group-hover:-translate-y-0.5" size={15} />
        </div>
      </button>
    </div>
  );
}

function Panel({ children, className = '' }: { children: ReactNode; className?: string }) {
  return <section className={`overflow-hidden pixel-panel ${className}`}>{children}</section>;
}

function PanelHead({
  eyebrow,
  title,
  action,
}: {
  eyebrow: string;
  title: string;
  action?: ReactNode;
}) {
  return (
    <div className="flex items-center justify-between gap-3 border-b-[3px] border-[#1a1a2e] bg-[#ffe566] px-4 py-3 sm:px-5">
      <div>
        <div className="pixel-font text-[8px] uppercase tracking-wider text-[#1a1a2e]/opacity-70">{eyebrow}</div>
        <h2 className="mt-1 text-[14px] font-bold tracking-tight text-[#1a1a2e]">{title}</h2>
      </div>
      {action}
    </div>
  );
}

function StepTracker({
  step,
  canMap,
  canReview,
  canExport,
  onStep,
}: {
  step: Step;
  canMap: boolean;
  canReview: boolean;
  canExport: boolean;
  onStep: (step: Step) => void;
}) {
  const allowed = {
    Upload: true,
    'Map fields': canMap,
    Review: canReview,
    Export: canExport,
  };
  const activeIndex = steps.indexOf(step);

  return (
    <nav className="pixel-panel px-4 py-3 sm:px-5" aria-label="Generation steps">
      <div className="grid grid-cols-4 gap-2">
        {steps.map((item, index) => (
          <div
            key={item}
            className={`flex min-w-0 items-center gap-2 text-[11px] font-bold ${
              item === step ? 'text-[#1a1a2e]' : index < activeIndex ? 'text-[#4ade80]' : 'text-[#1a1a2e]/40'
            }`}
          >
            <button
              type="button"
              disabled={!allowed[item]}
              onClick={() => onStep(item)}
              className="flex min-w-0 items-center gap-2 text-left"
              data-testid={`button-step-${normalize(item)}`}
            >
              <span
                className={`pixel-font grid h-7 w-7 shrink-0 place-items-center pixel-border-sm text-[10px] ${
                  item === step
                    ? 'bg-[#ffe566] text-[#1a1a2e]'
                    : index < activeIndex
                      ? 'bg-[#4ade80] text-[#1a1a2e]'
                      : 'bg-[#fff8e7] text-[#1a1a2e]/50'
                }`}
              >
                {index < activeIndex ? <Check size={12} /> : index + 1}
              </span>
              <span className="truncate">{item}</span>
            </button>
            {index < 3 && <span className="hidden h-[3px] flex-1 bg-[#1a1a2e]/20 sm:block" />}
          </div>
        ))}
      </div>
    </nav>
  );
}

function EmptyCanvas({ copy }: { copy: string }) {
  return (
    <div className="flex min-h-[315px] flex-col items-center justify-center pixel-grid px-8 text-center">
      <div className="mb-3 text-5xl leading-none select-none">🪪</div>
      <h3 className="pixel-font mb-2 text-[13px] text-[#1a1a2e]">CARD CANVAS</h3>
      <p className="max-w-xs text-[12px] leading-5 text-[#1a1a2e]/70">{copy}</p>
      <div className="pixel-btn mt-5 bg-[#4ade80] px-6 py-2 text-[11px] font-bold text-[#1a1a2e]">START</div>
    </div>
  );
}

function QualityRow({
  icon: Icon,
  title,
  detail,
  count,
  tone = 'green',
}: {
  icon: typeof Check;
  title: string;
  detail: string;
  count: string;
  tone?: 'green' | 'amber';
}) {
  return (
    <div className="flex items-center gap-2.5 pixel-border-sm bg-[#fff8e7] p-2.5" data-testid={`status-${normalize(title)}`}>
      <Icon size={16} className={tone === 'amber' ? 'text-[#e85d75]' : 'text-[#4ade80]'} />
      <div className="min-w-0">
        <strong className="block text-[11px] text-[#1a1a2e]">{title}</strong>
        <span className="block truncate text-[9px] text-[#1a1a2e]/60">{detail}</span>
      </div>
      <span className="pixel-font ml-auto text-[11px] text-[#1a1a2e]">{count}</span>
    </div>
  );
}

/* ========== MAIN STUDIO ========== */

function BulkIdStudio() {
  const [step, setStep] = useState<Step>('Upload');
  const [template, setTemplate] = useState<ParsedTemplate | null>(null);
  const [templateName, setTemplateName] = useState('');
  const [spreadsheet, setSpreadsheet] = useState<SpreadsheetData | null>(null);
  const [spreadsheetName, setSpreadsheetName] = useState('');
  const [archive, setArchive] = useState<PhotoArchive>();
  const [archiveName, setArchiveName] = useState('');
  const [mappings, setMappings] = useState<FieldMapping>({});
  const [selectedRow, setSelectedRow] = useState(0);
  const [search, setSearch] = useState('');
  const [previewUrl, setPreviewUrl] = useState('');
  const [generatedCards, setGeneratedCards] = useState<GeneratedCard[]>([]);
  const [generationIssues, setGenerationIssues] = useState<GenerationIssue[]>([]);
  const [isGenerating, setIsGenerating] = useState(false);
  const [progress, setProgress] = useState(0);
  const [notice, setNotice] = useState('');
  const [error, setError] = useState('');
  const [cardsPerPage, setCardsPerPage] = useState<4 | 9>(4);
  const [margin, setMargin] = useState(6);
  const [bleed, setBleed] = useState(0);
  const [photoCache, setPhotoCache] = useState<Record<string, string>>({});
  const [photoResolution, setPhotoResolution] = useState<Record<string, boolean>>({});
  const [photoResolutionReady, setPhotoResolutionReady] = useState(false);
  const [checkingPhotos, setCheckingPhotos] = useState(false);
  const [selectionMode, setSelectionMode] = useState(false);
  const [selectableSvg, setSelectableSvg] = useState('');
  const [selectedPhotoElementIndex, setSelectedPhotoElementIndex] = useState<number | null>(null);

  const photoCacheRef = useRef<Record<string, string>>({});
  const noticeTimer = useRef<number | undefined>(undefined);

  useEffect(() => {
    document.title = 'Bulk ID Studio — PIXEL';
    const description = document.querySelector('meta[name="description"]') ?? document.createElement('meta');
    description.setAttribute('name', 'description');
    description.setAttribute('content', 'Turn one SVG card template and a spreadsheet into press-ready identity cards in bulk.');
    document.head.appendChild(description);
  }, []);

  const flash = (message: string) => {
    setNotice(message);
    if (noticeTimer.current) window.clearTimeout(noticeTimer.current);
    noticeTimer.current = window.setTimeout(() => setNotice(''), 2800);
  };

  const handleTemplate = async (file: File) => {
    setError('');
    if (!file.name.toLowerCase().endsWith('.svg')) {
      setError('Please choose one SVG template file.');
      return;
    }
    try {
      const parsed = parseSvgTemplate(await file.text());
      setTemplate(parsed);
      setTemplateName(file.name);
      setSelectedRow(0);
      setSelectedPhotoElementIndex(null);
      setSelectionMode(false);
      setSelectableSvg('');
      setGeneratedCards([]);
      setGenerationIssues([]);
      setMappings(spreadsheet ? suggestMappings(parsed, spreadsheet) : {});
      if (spreadsheet) setStep('Map fields');
      flash(`${parsed.tokens.length} template field${parsed.tokens.length === 1 ? '' : 's'} detected.`);
    } catch {
      setError('This SVG could not be read. Try exporting it as plain SVG.');
    }
  };

  const handleSpreadsheet = async (file: File) => {
    setError('');
    if (!/\.(csv|xlsx|xls)$/i.test(file.name)) {
      setError('Please choose a CSV or XLSX spreadsheet.');
      return;
    }
    try {
      const parsed = await parseSpreadsheet(file);
      setSpreadsheet(parsed);
      setSpreadsheetName(file.name);
      setGeneratedCards([]);
      setGenerationIssues([]);
      if (template) {
        setMappings(suggestMappings(template, parsed));
        setStep('Map fields');
      }
      flash(`${parsed.rows.length} row${parsed.rows.length === 1 ? '' : 's'} imported.`);
    } catch {
      setError('The spreadsheet could not be parsed. Check that the first row contains headers.');
    }
  };

  const handleArchive = async (file: File) => {
    setError('');
    if (!file.name.toLowerCase().endsWith('.zip')) {
      setError('Photo matching accepts a ZIP of PNG, JPG, or WEBP images.');
      return;
    }
    try {
      const parsed = await parsePhotoArchive(file);
      setArchive(parsed);
      setArchiveName(file.name);
      setGeneratedCards([]);
      flash(`${parsed.names.length} photo${parsed.names.length === 1 ? '' : 's'} indexed from ZIP.`);
    } catch {
      setError('The photo ZIP could not be opened.');
    }
  };

  const rows = spreadsheet?.rows ?? [];
  const mappedCount = template?.tokens.filter((token) => Boolean(mappings[token] && mappings[token] !== '__ignore__')).length ?? 0;
  const photoToken = template?.tokens.find((token) => normalize(token) === 'photo');
  const photoHeader = photoToken ? mappings[photoToken] : '';

  useEffect(() => {
    if (!spreadsheet || !template?.photoPlaceholder || !photoToken || !photoHeader) {
      setPhotoResolution({});
      setPhotoResolutionReady(!template?.photoPlaceholder);
      setCheckingPhotos(false);
      return;
    }
    let cancelled = false;
    setCheckingPhotos(true);
    setPhotoResolutionReady(false);

    const checkPhotos = async () => {
      const results = await Promise.all(
        spreadsheet.rows.map(async (row, index) => {
          const value = String(row[photoHeader] ?? '').trim();
          const key = `${index}:${value}`;
          const embedded = spreadsheet.embeddedPhotos?.[index];
          if (embedded) return [key, true] as const;
          if (!value) return [key, false] as const;
          const resolved = await resolvePhotoData(value, archive);
          return [key, Boolean(resolved)] as const;
        }),
      );
      if (cancelled) return;
      setPhotoResolution(Object.fromEntries(results));
      setPhotoResolutionReady(true);
      setCheckingPhotos(false);
    };
    void checkPhotos();
    return () => { cancelled = true; };
  }, [spreadsheet, template, photoToken, photoHeader, archive]);

  const photoIsResolved = (row: SpreadsheetRow, index: number) => {
    if (spreadsheet?.embeddedPhotos?.[index]) return true;
    const value = String(photoHeader ? row[photoHeader] ?? '' : '').trim();
    if (!value) return false;
    if (archive?.files[value.toLowerCase()] || archive?.files[value.split('/').pop()?.toLowerCase() ?? '']) return true;
    return photoResolution[`${index}:${value}`] === true;
  };

  const qualityIssues = useMemo(() => {
    if (!spreadsheet || !template) return [];
    const issues: GenerationIssue[] = [];
    spreadsheet.rows.forEach((row, index) => {
      template.tokens.forEach((token) => {
        const header = mappings[token];
        if (token !== photoToken && (!header || header === '__ignore__' || !String(row[header] ?? '').trim())) {
          issues.push({
            row: index + 1,
            field: token,
            message: header ? `${token} is empty` : `${token} has no mapped column`,
            person: rowLabel(row, mappings, index),
          });
        }
      });
      if (photoToken && template.photoPlaceholder) {
        const value = photoHeader ? row[photoHeader] ?? '' : '';
        const resolved = photoIsResolved(row, index);
        const embedded = spreadsheet.embeddedPhotos?.[index];
        if (!embedded && (!photoHeader || !value)) {
          issues.push({
            row: index + 1,
            field: 'Photo',
            message: photoHeader ? 'photo value is empty' : 'Photo has no mapped column',
            person: rowLabel(row, mappings, index),
          });
        } else if (!embedded && photoResolutionReady && !resolved) {
          issues.push({
            row: index + 1,
            field: 'Photo',
            message: 'photo could not be matched',
            person: rowLabel(row, mappings, index),
          });
        }
      }
    });
    return issues;
  }, [spreadsheet, template, mappings, photoToken, photoHeader, archive, photoResolution, photoResolutionReady]);

  const issues = useMemo(() => {
    const all = [...qualityIssues, ...generationIssues];
    return all.filter((issue, index, list) =>
      list.findIndex((c) => c.row === issue.row && c.field === issue.field && c.message === issue.message) === index,
    );
  }, [qualityIssues, generationIssues]);

  const matchedPhotoCount = useMemo(() => {
    if (!spreadsheet || !photoHeader) return 0;
    return spreadsheet.rows.filter((row, index) => photoIsResolved(row, index)).length;
  }, [spreadsheet, photoHeader, archive, photoResolution, photoCache]);

  const validIndexes = useMemo(
    () => rows.map((_, index) => index).filter((index) => !qualityIssues.some((issue) => issue.row === index + 1)),
    [rows, qualityIssues],
  );

  const filteredIndexes = useMemo(
    () => validIndexes.filter((index) => rowLabel(rows[index], mappings, index).toLowerCase().includes(search.toLowerCase())),
    [validIndexes, rows, mappings, search],
  );

  const activeRow = rows[selectedRow];
  const hasMappings = Boolean(template && spreadsheet && template.tokens.every((token) => mappings[token] && mappings[token] !== '__ignore__'));
  const canMap = Boolean(template && spreadsheet);
  const canExport = generatedCards.length > 0;
  const canGenerate = Boolean(canMap && hasMappings && rows.length > 0 && !checkingPhotos);
  const canReview = canExport;

  useEffect(() => {
    if (!activeRow || !template) {
      setPreviewUrl('');
      setSelectableSvg('');
      return;
    }
    let cancelled = false;
    const draw = async () => {
      try {
        const photoValue = photoToken && photoHeader ? activeRow[photoHeader] ?? '' : '';
        const cacheKey = `${selectedRow}:${photoValue}`;
        let photoData: string | undefined = photoCacheRef.current[cacheKey];
        if (!photoData) {
          const embedded = spreadsheet?.embeddedPhotos?.[selectedRow];
          if (embedded) photoData = await fileToDataUrl(embedded);
        }
        if (!photoData && photoValue) {
          photoData = await resolvePhotoData(photoValue, archive);
          if (photoData) {
            photoCacheRef.current[cacheKey] = photoData;
            setPhotoCache((c) => ({ ...c, [cacheKey]: photoData as string }));
          }
        }

        const selectableSource = template.rawSvg.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (_match, rawToken: string) => {
          const token = rawToken.trim();
          if (token.toLowerCase() === 'image' || token.toLowerCase() === 'photo') return '';
          const header = mappings[token];
          return header && header !== '__ignore__' ? escapeSvgText(String(activeRow[header] ?? '')) : '';
        });

        const selectable = makeSelectableSvg(selectableSource, selectedPhotoElementIndex);
        if (!cancelled) setSelectableSvg(selectable);

        const svg = renderSvgForRow(template, activeRow, mappings, photoData, selectedPhotoElementIndex ?? undefined);
        const blob = await renderSvgToPng(svg, 2);
        if (!cancelled) setPreviewUrl(URL.createObjectURL(blob));
      } catch {
        if (!cancelled) {
          setPreviewUrl('');
          setSelectableSvg('');
        }
      }
    };
    void draw();
    return () => { cancelled = true; };
  }, [activeRow, template, mappings, photoToken, photoHeader, archive, selectedRow, selectedPhotoElementIndex, spreadsheet]);

  useEffect(() => () => {
    if (previewUrl) URL.revokeObjectURL(previewUrl);
  }, [previewUrl]);

  const handleSvgElementClick = (event: MouseEvent<HTMLDivElement>) => {
    if (!selectionMode) return;
    const target = event.target as Element | null;
    if (!target) return;
    const element = target.closest('[data-bulk-id-element]');
    if (!element) return;
    const rawIndex = element.getAttribute('data-bulk-id-element');
    const index = Number(rawIndex);
    if (!Number.isInteger(index)) return;
    setSelectedPhotoElementIndex(index);
    if (template) setTemplate(setPhotoPlaceholder(template, index));
    setGeneratedCards([]);
    setGenerationIssues([]);
    flash('Photo area selected ✓ Click "Use this area" to confirm.');
  };

  const updateMapping = (token: string, header: string) => {
    setMappings((c) => ({ ...c, [token]: header }));
    setGeneratedCards([]);
    setGenerationIssues([]);
    flash(`${token} mapped to ${header}.`);
  };

  const renderCard = async (rowIndex: number): Promise<GeneratedCard | null> => {
    if (!template || !spreadsheet) return null;
    const row = spreadsheet.rows[rowIndex];
    const photoValue = photoToken && photoHeader ? row[photoHeader] ?? '' : '';
    const cacheKey = `${rowIndex}:${photoValue}`;
    let photoData: string | undefined = photoCacheRef.current[cacheKey];
    if (!photoData) {
      const embedded = spreadsheet.embeddedPhotos?.[rowIndex];
      if (embedded) photoData = await fileToDataUrl(embedded);
    }
    if (!photoData && photoValue) {
      photoData = await resolvePhotoData(photoValue, archive);
      if (photoData) {
        photoCacheRef.current[cacheKey] = photoData;
        setPhotoCache((c) => ({ ...c, [cacheKey]: photoData as string }));
      }
    }
    if (photoToken && template.photoPlaceholder && !photoData) {
      setGenerationIssues((c) => [...c, {
        row: rowIndex + 1,
        field: 'Photo',
        message: 'photo could not be resolved during rendering',
        person: rowLabel(row, mappings, rowIndex),
      }]);
      return null;
    }
    const svgText = renderSvgForRow(template, row, mappings, photoData, selectedPhotoElementIndex ?? undefined);
    const pngBlob = await renderSvgToPng(svgText, 3);
    return { rowIndex, pngBlob, svgText };
  };

  const generate = async (allRows: boolean) => {
    if (!canGenerate || isGenerating) return;
    if (!validIndexes.length) {
      setError('There are no valid rows yet. Resolve the quality issues before generating.');
      return;
    }
    setError('');
    setIsGenerating(true);
    setProgress(0);
    setGenerationIssues([]);
    const targets = allRows ? validIndexes : validIndexes.slice(0, 5);
    const result: GeneratedCard[] = [];
    for (let index = 0; index < targets.length; index += 1) {
      const card = await renderCard(targets[index]);
      if (card) result.push(card);
      setProgress(Math.round(((index + 1) / targets.length) * 100));
      await new Promise<void>((r) => window.setTimeout(r, 0));
    }
    setGeneratedCards(result);
    setIsGenerating(false);
    if (result.length) {
      setStep(allRows ? 'Export' : 'Review');
      flash(allRows ? `${result.length} cards rendered and ready to export.` : `${result.length} preview cards rendered.`);
    } else {
      setError('No cards could be rendered. Resolve the issues in the quality panel and try again.');
    }
  };

  const exportZip = async () => {
    if (!generatedCards.length) return;
    const zip = new JSZip();
    generatedCards.forEach((card) => {
      const number = String(card.rowIndex + 1).padStart(4, '0');
      zip.file(`card-${number}.png`, card.pngBlob);
      zip.file(`card-${number}.svg`, card.svgText);
    });
    const blob = await zip.generateAsync({ type: 'blob' });
    downloadBlob(blob, 'bulk-id-studio-cards.zip');
    flash('Card ZIP downloaded.');
  };

  const exportPdf = async () => {
    if (!generatedCards.length) return;
    const pdf = await PDFDocument.create();
    const pageWidth = 612;
    const pageHeight = 792;
    const mm = (v: number) => v * 2.83465;
    const marginPt = mm(margin);
    const bleedPt = mm(bleed);
    const columns = cardsPerPage === 4 ? 2 : 3;
    const rowsOnPage = cardsPerPage === 4 ? 2 : 3;
    const gap = mm(5);
    const cr80Ratio = 85.6 / 53.98;
    const availableWidth = pageWidth - marginPt * 2 - gap * (columns - 1);
    const cardWidth = availableWidth / columns;
    const cardHeight = cardWidth / cr80Ratio;
    const totalHeight = cardHeight * rowsOnPage + gap * (rowsOnPage - 1);

    for (let offset = 0; offset < generatedCards.length; offset += cardsPerPage) {
      const page = pdf.addPage([pageWidth, pageHeight]);
      const pageCards = generatedCards.slice(offset, offset + cardsPerPage);
      await Promise.all(
        pageCards.map(async (card, index) => {
          const col = index % columns;
          const row = Math.floor(index / columns);
          const image = await pdf.embedPng(await card.pngBlob.arrayBuffer());
          page.drawImage(image, {
            x: marginPt + col * (cardWidth + gap) - bleedPt,
            y: pageHeight - marginPt - totalHeight + (rowsOnPage - 1 - row) * (cardHeight + gap) - bleedPt,
            width: cardWidth + bleedPt * 2,
            height: cardHeight + bleedPt * 2,
          });
        }),
      );
      await new Promise<void>((r) => window.setTimeout(r, 0));
    }
    const pdfBytes = await pdf.save();
    const pdfBuffer = new ArrayBuffer(pdfBytes.byteLength);
    new Uint8Array(pdfBuffer).set(pdfBytes);
    const blob = new Blob([pdfBuffer], { type: 'application/pdf' });
    downloadBlob(blob, `bulk-id-studio-${cardsPerPage}up.pdf`);
    flash('Print PDF downloaded.');
  };

  const resetStudio = () => {
    setTemplate(null);
    setTemplateName('');
    setSpreadsheet(null);
    setSpreadsheetName('');
    setArchive(undefined);
    setArchiveName('');
    setMappings({});
    setGeneratedCards([]);
    setGenerationIssues([]);
    setSelectedRow(0);
    setStep('Upload');
    setError('');
    setProgress(0);
    photoCacheRef.current = {};
    setPhotoCache({});
    setSelectionMode(false);
    setSelectableSvg('');
    setSelectedPhotoElementIndex(null);
    flash('Studio cleared.');
  };

  return (
    <main
      className="min-h-[100dvh] px-3 py-5 text-[#1a1a2e] sm:px-6 sm:py-7 lg:px-8"
      style={{
        backgroundImage: 'linear-gradient(180deg, #7ec8f5 0%, #b8e0f8 35%, #e8f4fc 65%, #f0e6c8 100%)',
      }}
    >
      <div className="mx-auto max-w-[1480px]">
        {/* HEADER */}
        <header className="flex flex-col justify-between gap-5 pb-5 md:flex-row md:items-start">
          <div>
            <div className="pixel-font mb-2 text-[9px] uppercase tracking-widest text-[#1a1a2e]/60">
              BULK ID STUDIO / PIXEL WORKSPACE
            </div>
            <h1 className="text-[26px] font-bold leading-tight sm:text-[30px]">
              One template.<br />
              <span className="text-[#e85d75]">Every card.</span>
            </h1>
            <p className="mt-2 max-w-xl text-[12px] leading-5 text-[#1a1a2e]/70">
              A quiet, dependable path from SVG and spreadsheet to press-ready identity cards.
            </p>
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              className="pixel-btn inline-flex items-center gap-2 bg-[#fff8e7] px-3 py-2 text-[10px] font-bold text-[#1a1a2e]"
              onClick={() => flash('Your batch stays in this browser until you download it.')}
              data-testid="button-autosave"
            >
              <CloudUpload size={14} />
              Local workspace
            </button>
            <button
              type="button"
              aria-label="Studio settings"
              className="pixel-btn grid h-9 w-9 place-items-center bg-[#fff8e7] text-[#1a1a2e]"
              onClick={() => flash('Print settings are available in Export.')}
              data-testid="button-settings"
            >
              <Settings2 size={15} />
            </button>
            <button
              type="button"
              aria-label="Help"
              className="pixel-btn grid h-9 w-9 place-items-center bg-[#fff8e7] text-[#1a1a2e]"
              onClick={() => flash('Use {{Token}} text in your SVG, then map each token to a column.')}
              data-testid="button-help"
            >
              <CircleHelp size={15} />
            </button>
          </div>
        </header>

        <StepTracker step={step} canMap={canMap} canReview={canReview} canExport={canExport} onStep={setStep} />

        {(error || notice) && (
          <div
            className={`mt-3 flex items-center gap-2 pixel-border-sm px-3 py-2.5 text-[11px] ${
              error ? 'bg-[#ff6b6b] text-white' : 'bg-[#4ade80] text-[#1a1a2e]'
            }`}
            role="status"
            data-testid={error ? 'status-error' : 'status-notice'}
          >
            {error ? <AlertTriangle size={14} /> : <Check size={14} />}
            {error || notice}
            {error && (
              <button type="button" className="ml-auto" onClick={() => setError('')} aria-label="Dismiss error" data-testid="button-dismiss-error">
                <X size={14} />
              </button>
            )}
          </div>
        )}

        <section className="mt-4 grid items-start gap-3 lg:grid-cols-[290px_minmax(400px,1fr)_320px]">
          {/* LEFT */}
          <aside className="space-y-3">
            <Panel>
              <PanelHead
                eyebrow="Source files"
                title="Batch ingredients"
                action={
                  <button
                    type="button"
                    aria-label="Clear workspace"
                    className="pixel-btn grid h-7 w-7 place-items-center bg-[#fff8e7] text-[#1a1a2e]"
                    onClick={resetStudio}
                    data-testid="button-clear-workspace"
                  >
                    <MoreHorizontal size={14} />
                  </button>
                }
              />
              <div className="space-y-2.5 p-4 sm:p-5">
                <UploadZone kind="template" accept=".svg,image/svg+xml" icon={FileText} title="Add SVG template" detail="SVG only · one card face" onFile={handleTemplate} fileName={templateName} />
                <UploadZone kind="spreadsheet" accept=".csv,.xlsx,.xls,text/csv" icon={FileSpreadsheet} title="Add spreadsheet" detail="CSV or XLSX · first row = headers" onFile={handleSpreadsheet} fileName={spreadsheetName} />
                <UploadZone kind="photos" accept=".zip,application/zip" icon={FileArchive} title="Add photo ZIP" detail="Optional · PNG, JPG, WEBP" onFile={handleArchive} fileName={archiveName} />

                {template && (
                  <div className="border-t-[3px] border-[#1a1a2e] pt-4">
                    <div className="pixel-font text-[8px] uppercase tracking-wider text-[#1a1a2e]/60">
                      Detected fields <span className="text-[#4ade80]">· {template.tokens.length}</span>
                    </div>
                    {template.tokens.length ? (
                      <div className="mt-2 flex flex-wrap gap-1.5">
                        {template.tokens.map((token) => (
                          <span key={token} className="pixel-font pixel-border-sm bg-[#ffe566] px-2 py-1 text-[9px] text-[#1a1a2e]" data-testid={`text-token-${normalize(token)}`}>
                            {`{{${token}}}`}
                          </span>
                        ))}
                      </div>
                    ) : (
                      <p className="mt-2 text-[10px] text-[#e85d75]">No {'{{Token}}'} fields found in this SVG.</p>
                    )}
                    {template.photoPlaceholder && (
                      <div className="mt-2 flex items-center gap-1.5 text-[10px] text-[#4ade80]">
                        <ImageIcon size={13} /> Photo bounds detected
                      </div>
                    )}
                  </div>
                )}

                {!template && !spreadsheet && (
                  <div className="flex gap-2 pixel-border-sm bg-[#ffe566] p-2.5 text-[10px] leading-4 text-[#1a1a2e]">
                    <Sparkles className="mt-0.5 shrink-0" size={13} />
                    Start with an SVG and spreadsheet. Your files are parsed locally in this browser.
                  </div>
                )}
              </div>
            </Panel>

            {step === 'Upload' && (
              <Panel className="hidden lg:block">
                <div className="p-4">
                  <div className="pixel-font text-[8px] uppercase tracking-wider text-[#1a1a2e]/60">Workflow note</div>
                  <p className="mt-2 text-[11px] leading-5 text-[#1a1a2e]/70">
                    Nothing is generated until every field can be traced back to a source column.
                  </p>
                </div>
              </Panel>
            )}
          </aside>

          {/* CENTER */}
          <section className="min-w-0">
            <Panel className="min-h-[580px]">
              <PanelHead
                eyebrow={`${step === 'Upload' ? 'Step 01' : step === 'Map fields' ? 'Step 02' : step === 'Review' ? 'Step 03' : 'Step 04'} / live workspace`}
                title={
                  step === 'Upload' ? 'Bring in your source files'
                    : step === 'Map fields' ? 'Map fields & preview cards'
                      : step === 'Review' ? 'Review generated cards'
                        : 'Package your press files'
                }
                action={
                  template && spreadsheet ? (
                    <div className="flex items-center gap-2">
                      <div className="hidden items-center gap-1.5 pixel-border-sm bg-[#fff8e7] px-2 py-1.5 text-[#1a1a2e]/50 sm:flex">
                        <Search size={12} />
                        <input
                          className="w-28 bg-transparent text-[10px] text-[#1a1a2e] outline-none"
                          placeholder="Find a row…"
                          value={search}
                          onChange={(e) => setSearch(e.target.value)}
                          data-testid="input-search-rows"
                        />
                      </div>
                      <button
                        type="button"
                        className="pixel-btn grid h-8 w-8 place-items-center bg-[#fff8e7] text-[#1a1a2e]"
                        onClick={() => { setGeneratedCards([]); flash('Preview refreshed.'); }}
                        aria-label="Refresh preview"
                        data-testid="button-refresh-preview"
                      >
                        <RefreshCw size={13} />
                      </button>
                    </div>
                  ) : undefined
                }
              />

              {!template || !spreadsheet ? (
                <EmptyCanvas copy="Upload the SVG face and the spreadsheet on the left. Once both are present, the live canvas will use your actual row data." />
              ) : !hasMappings ? (
                <div className="flex min-h-[315px] flex-col items-center justify-center pixel-grid px-8 text-center" data-testid="empty-mapping-state">
                  <div className="text-4xl mb-3">🗺️</div>
                  <h3 className="pixel-font text-[13px] text-[#1a1a2e] mb-2">MAP YOUR FIELDS</h3>
                  <p className="max-w-xs text-[12px] leading-5 text-[#1a1a2e]/70">
                    Choose a spreadsheet column for every detected token. The preview updates as soon as the mapping is complete.
                  </p>
                  <button
                    type="button"
                    className="pixel-btn mt-4 inline-flex items-center gap-2 bg-[#ffe566] px-4 py-2 text-[11px] font-bold text-[#1a1a2e]"
                    onClick={() => setStep('Map fields')}
                    data-testid="button-open-mapping"
                  >
                    <ArrowRight size={13} /> Open field mapping
                  </button>
                </div>
              ) : (
                <>
                  <div className="flex flex-wrap items-center justify-between gap-2 border-b-[3px] border-[#1a1a2e] bg-[#fff8e7] px-4 py-2.5 sm:px-5">
                    <div>
                      <div className="pixel-font text-[8px] uppercase tracking-wider text-[#1a1a2e]/60">Photo placement</div>
                      <p className="mt-0.5 text-[10px] text-[#1a1a2e]/70">
                        {selectionMode
                          ? 'Click the exact shape or image that should hold the photo.'
                          : selectedPhotoElementIndex !== null
                            ? 'Photo area selected and locked for generation.'
                            : 'Want a custom photo area? Pick it directly on the SVG.'}
                      </p>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        className={`pixel-btn inline-flex items-center justify-center gap-1.5 px-3 py-2 text-[10px] font-bold ${
                          selectionMode ? 'bg-[#e85d75] text-white' : 'bg-[#fff8e7] text-[#1a1a2e]'
                        }`}
                        onClick={() => {
                          setSelectionMode((c) => !c);
                          if (!selectionMode) flash('Selection mode on — click the exact photo area in the SVG.');
                        }}
                      >
                        <ImageIcon size={13} />
                        {selectionMode ? 'Click an element' : selectedPhotoElementIndex !== null ? 'Change photo area' : 'Choose photo area'}
                      </button>
                      {selectionMode && selectedPhotoElementIndex !== null && (
                        <button
                          type="button"
                          className="pixel-btn inline-flex items-center justify-center gap-1.5 bg-[#4ade80] px-3 py-2 text-[10px] font-bold text-[#1a1a2e]"
                          onClick={() => { setSelectionMode(false); flash('Photo area confirmed ✓'); }}
                        >
                          <Check size={13} /> Use this area
                        </button>
                      )}
                    </div>
                  </div>

                  {selectionMode && selectedPhotoElementIndex !== null && (
                    <div className="border-b-[3px] border-[#1a1a2e] bg-[#4ade80] px-4 py-2 text-[10px] font-bold text-[#1a1a2e] sm:px-5">
                      ✓ Photo area selected. Click <strong className="mx-1">Use this area</strong> when you&apos;re happy with it.
                    </div>
                  )}

                  {selectionMode && selectableSvg ? (
                    <div className="relative flex min-h-[315px] items-center justify-center overflow-hidden pixel-grid p-6">
                      <span className="pixel-font absolute left-3 top-3 z-10 pixel-border-sm bg-[#ffe566] px-2 py-1 text-[8px] text-[#1a1a2e]">
                        CLICK TO SELECT PHOTO AREA
                      </span>
                      <div
                        className="flex max-h-[430px] max-w-[78%] items-center justify-center pixel-border bg-[#fff8e7] p-2"
                        onClick={handleSvgElementClick}
                        dangerouslySetInnerHTML={{ __html: selectableSvg }}
                        data-testid="svg-selection-canvas"
                      />
                    </div>
                  ) : previewUrl ? (
                    <div className="relative flex min-h-[315px] items-center justify-center overflow-hidden pixel-grid p-6">
                      <span className="pixel-font absolute left-3 top-3 text-[8px] text-[#1a1a2e]/60">
                        SVG CANVAS / ROW {selectedRow + 1}
                      </span>
                      <img
                        src={previewUrl}
                        alt={`Rendered card for row ${selectedRow + 1}`}
                        className="max-h-[390px] max-w-[72%] pixel-border bg-[#fff8e7]"
                        data-testid="img-live-preview"
                      />
                    </div>
                  ) : (
                    <div className="flex min-h-[315px] items-center justify-center pixel-grid">
                      <LoaderCircle className="spin text-[#1a1a2e]" size={24} />
                    </div>
                  )}
                </>
              )}

              {template && spreadsheet && hasMappings && (
                <div className="flex flex-wrap items-center justify-between gap-2 border-b-[3px] border-[#1a1a2e] px-4 py-3 text-[10px] text-[#1a1a2e]/70 sm:px-5">
                  <span>
                    Showing <strong className="text-[#1a1a2e]" data-testid="text-active-person">
                      {activeRow ? rowLabel(activeRow, mappings, selectedRow) : '—'}
                    </strong>{' '}
                    · row <strong className="text-[#1a1a2e]">{Math.min(selectedRow + 1, rows.length)}</strong> of {rows.length}
                  </span>
                  <div className="flex items-center gap-2">
                    <span className="pixel-font hidden text-[8px] sm:inline">CR80 · 85.6 × 53.98 mm</span>
                    <button
                      type="button"
                      className="pixel-btn grid h-7 w-7 place-items-center bg-[#fff8e7] text-[#1a1a2e]"
                      disabled={selectedRow <= 0}
                      onClick={() => setSelectedRow((c) => Math.max(0, c - 1))}
                      aria-label="Previous row"
                      data-testid="button-previous-row"
                    >
                      <ArrowLeft size={12} />
                    </button>
                    <button
                      type="button"
                      className="pixel-btn grid h-7 w-7 place-items-center bg-[#fff8e7] text-[#1a1a2e]"
                      disabled={selectedRow >= rows.length - 1}
                      onClick={() => setSelectedRow((c) => Math.min(rows.length - 1, c + 1))}
                      aria-label="Next row"
                      data-testid="button-next-row"
                    >
                      <ArrowRight size={12} />
                    </button>
                  </div>
                </div>
              )}

              {template && spreadsheet && (
                <div className="border-b-[3px] border-[#1a1a2e] px-4 py-3 sm:px-5">
                  <div className="flex items-center justify-between">
                    <div>
                      <div className="pixel-font text-[8px] uppercase tracking-wider text-[#1a1a2e]/60">Rendered gallery</div>
                      <h3 className="mt-1 text-[13px] font-bold text-[#1a1a2e]">
                        {generatedCards.length
                          ? `${generatedCards.length} actual PNG${generatedCards.length === 1 ? '' : 's'} generated`
                          : 'First five valid rows'}
                      </h3>
                    </div>
                    <Pill tone={generatedCards.length ? 'green' : 'neutral'}>
                      {generatedCards.length ? `${generatedCards.length} ready` : `${validIndexes.length} valid`}
                    </Pill>
                  </div>

                  {generatedCards.length ? (
                    <div className="mt-3 grid grid-cols-2 gap-2 sm:grid-cols-5">
                      {generatedCards.slice(0, 10).map((card) => (
                        <GalleryCard key={card.rowIndex} card={card} selected={card.rowIndex === selectedRow} onClick={() => setSelectedRow(card.rowIndex)} />
                      ))}
                    </div>
                  ) : (
                    <div className="mt-3 grid grid-cols-2 gap-2 sm:grid-cols-5">
                      {filteredIndexes.slice(0, 5).map((index) => (
                        <button
                          type="button"
                          key={index}
                          className={`pixel-border-sm bg-[#fff8e7] p-2 text-left transition-transform hover:-translate-y-0.5 ${
                            selectedRow === index ? 'ring-2 ring-[#ffe566]' : ''
                          }`}
                          onClick={() => setSelectedRow(index)}
                          data-testid={`button-row-${index + 1}`}
                        >
                          <div className="flex h-24 items-center justify-center bg-[#7ec8f5]/30 text-[#1a1a2e]">
                            <ImageIcon size={20} strokeWidth={1.5} />
                          </div>
                          <span className="mt-1.5 block truncate text-[9px] font-bold text-[#1a1a2e]">
                            {rowLabel(rows[index], mappings, index)}
                          </span>
                          <span className="pixel-font block text-[8px] text-[#1a1a2e]/50">row {index + 1}</span>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </Panel>
          </section>

          {/* RIGHT */}
          <aside className="space-y-3">
            <Panel>
              <PanelHead
                eyebrow="Quality check"
                title="Before you generate"
                action={<ShieldCheck size={17} className={issues.length ? 'text-[#e85d75]' : 'text-[#4ade80]'} />}
              />
              <div className="space-y-2 p-4 sm:p-5">
                <QualityRow
                  icon={issues.some((i) => i.field !== 'Photo') ? AlertTriangle : Check}
                  title="Fields mapped"
                  detail={template ? `${mappedCount} of ${template.tokens.length} tokens have a source` : 'Upload a template first'}
                  count={template ? `${mappedCount}/${template.tokens.length}` : '—'}
                  tone={issues.some((i) => i.field !== 'Photo') ? 'amber' : 'green'}
                />
                <QualityRow
                  icon={ImageIcon}
                  title="Photos matched"
                  detail={
                    template?.photoPlaceholder
                      ? checkingPhotos ? 'Checking image sources…' : `${matchedPhotoCount} of ${rows.length} rows ready`
                      : 'No photo placeholder detected'
                  }
                  count={template?.photoPlaceholder ? `${matchedPhotoCount}/${rows.length}` : '—'}
                  tone={issues.some((i) => i.field === 'Photo') ? 'amber' : 'green'}
                />
                <QualityRow
                  icon={Layers3}
                  title="Rows imported"
                  detail={spreadsheet ? 'Spreadsheet is ready' : 'Waiting for spreadsheet'}
                  count={spreadsheet ? String(rows.length) : '—'}
                />

                {issues.length > 0 && (
                  <div className="pixel-border-sm bg-[#ffe566] p-3" data-testid="quality-issues">
                    <div className="mb-2 flex items-center justify-between text-[#1a1a2e]">
                      <strong className="text-[11px]">{issues.length} item{issues.length === 1 ? '' : 's'} to resolve</strong>
                      <AlertTriangle size={14} />
                    </div>
                    <ul className="m-0 grid max-h-44 gap-2 overflow-auto p-0">
                      {issues.slice(0, 20).map((issue, index) => (
                        <li key={`${issue.row}-${issue.field}-${index}`} className="flex gap-1.5 text-[10px] leading-4 text-[#1a1a2e]/80">
                          <AlertTriangle className="mt-0.5 shrink-0" size={11} />
                          <span>
                            <strong>Row {issue.row} · {issue.person}</strong> — {issue.message}
                          </span>
                        </li>
                      ))}
                    </ul>
                  </div>
                )}

                {template && spreadsheet && !issues.length && (
                  <div className="flex gap-2 pixel-border-sm bg-[#4ade80] p-2.5 text-[10px] leading-4 text-[#1a1a2e]" data-testid="status-quality-ready">
                    <Check className="mt-0.5 shrink-0" size={13} />
                    Every imported row is ready for rendering.
                  </div>
                )}
              </div>
            </Panel>

            <Panel>
              <PanelHead
                eyebrow="Field mapping"
                title="Spreadsheet → template"
                action={hasMappings ? <Pill tone="green">Synced</Pill> : <Pill>Required</Pill>}
              />
              <div className="p-4 sm:p-5">
                {template?.tokens.length ? (
                  template.tokens.map((token) => (
                    <div key={token} className="grid grid-cols-[1fr_15px_1.18fr] items-center gap-1.5 border-b border-[#1a1a2e]/15 py-2 last:border-0">
                      <code className="pixel-font truncate text-[10px] text-[#1a1a2e]">{`{{${token}}}`}</code>
                      <ArrowRight size={12} className="text-[#1a1a2e]/40" />
                      <select
                        className="min-w-0 pixel-border-sm bg-[#fff8e7] px-1.5 py-1.5 text-[10px] text-[#1a1a2e] outline-none"
                        value={mappings[token] ?? ''}
                        onChange={(e) => updateMapping(token, e.target.value)}
                        disabled={!spreadsheet}
                        aria-label={`Map ${token}`}
                        data-testid={`select-mapping-${normalize(token)}`}
                      >
                        <option value="">Choose column</option>
                        <option value="__ignore__">Ignore column</option>
                        {spreadsheet?.headers.map((header) => (
                          <option value={header} key={header}>{header}</option>
                        ))}
                      </select>
                    </div>
                  ))
                ) : (
                  <p className="text-[10px] leading-4 text-[#1a1a2e]/60">Detected {'{{Token}}'} elements will appear here.</p>
                )}
              </div>
            </Panel>

            <Panel>
              <div className="p-4 sm:p-5">
                <div className="flex items-center gap-2">
                  <Zap size={15} className="text-[#e85d75]" />
                  <h2 className="text-[13px] font-bold text-[#1a1a2e]">
                    {canExport ? 'Batch is rendered' : 'Ready when you are'}
                  </h2>
                </div>
                <p className="mt-2 text-[10px] leading-4 text-[#1a1a2e]/70">
                  {canExport
                    ? 'Download the generated cards or lay them out for print.'
                    : 'Generate five valid rows first. This catches layout and image issues before the full run.'}
                </p>

                {isGenerating && (
                  <div className="mt-3">
                    <div className="mb-1 flex justify-between pixel-font text-[8px] text-[#1a1a2e]">
                      <span>Rendering in browser</span>
                      <span>{progress}%</span>
                    </div>
                    <div className="h-2 overflow-hidden pixel-border-sm bg-[#fff8e7]">
                      <div className="h-full bg-[#4ade80] transition-[width] duration-200" style={{ width: `${progress}%` }} />
                    </div>
                  </div>
                )}

                <div className="mt-3 grid gap-2">
                  {!canExport ? (
                    <button
                      type="button"
                      className="pixel-btn inline-flex items-center justify-center gap-2 bg-[#ffe566] px-3 py-2.5 text-[11px] font-bold text-[#1a1a2e] disabled:opacity-50"
                      disabled={!canGenerate || isGenerating}
                      onClick={() => void generate(false)}
                      data-testid="button-generate-preview"
                    >
                      {isGenerating ? (
                        <><LoaderCircle className="spin" size={14} /> Rendering {progress}%</>
                      ) : (
                        <><Play size={14} /> Generate preview batch</>
                      )}
                    </button>
                  ) : (
                    <>
                      <button
                        type="button"
                        className="pixel-btn inline-flex items-center justify-center gap-2 bg-[#ffe566] px-3 py-2.5 text-[11px] font-bold text-[#1a1a2e] disabled:opacity-50"
                        disabled={isGenerating || validIndexes.length === 0}
                        onClick={() => void generate(true)}
                        data-testid="button-generate-all"
                      >
                        {isGenerating ? (
                          <><LoaderCircle className="spin" size={14} /> Rendering {progress}%</>
                        ) : (
                          <><RefreshCw size={14} /> Generate all · {validIndexes.length}</>
                        )}
                      </button>
                      <button
                        type="button"
                        className="pixel-btn inline-flex items-center justify-center gap-2 bg-[#fff8e7] px-3 py-2 text-[10px] font-bold text-[#1a1a2e]"
                        onClick={() => void exportZip()}
                        data-testid="button-export-zip"
                      >
                        <Download size={13} /> Download PNG + SVG ZIP
                      </button>
                    </>
                  )}
                </div>
              </div>
            </Panel>
          </aside>
        </section>

        {step === 'Export' && canExport && (
          <section className="mt-3 grid gap-3 lg:grid-cols-[1fr_320px]">
            <Panel>
              <PanelHead eyebrow="Print output" title="Build a combined PDF" action={<Printer size={16} className="text-[#4ade80]" />} />
              <div className="grid gap-4 p-4 sm:grid-cols-3 sm:p-5">
                <label className="text-[10px] font-bold text-[#1a1a2e]">
                  Cards per page
                  <select
                    className="mt-1.5 block w-full pixel-border-sm bg-[#fff8e7] px-2 py-2 text-[11px] text-[#1a1a2e] outline-none"
                    value={cardsPerPage}
                    onChange={(e) => setCardsPerPage(Number(e.target.value) as 4 | 9)}
                    data-testid="select-cards-per-page"
                  >
                    <option value="4">4-up · larger cards</option>
                    <option value="9">9-up · compact sheet</option>
                  </select>
                </label>
                <label className="text-[10px] font-bold text-[#1a1a2e]">
                  Margin · mm
                  <input
                    className="mt-1.5 block w-full pixel-border-sm bg-[#fff8e7] px-2 py-2 text-[11px] text-[#1a1a2e] outline-none"
                    type="number" min="0" max="30" value={margin}
                    onChange={(e) => setMargin(Number(e.target.value))}
                    data-testid="input-print-margin"
                  />
                </label>
                <label className="text-[10px] font-bold text-[#1a1a2e]">
                  Bleed · mm
                  <input
                    className="mt-1.5 block w-full pixel-border-sm bg-[#fff8e7] px-2 py-2 text-[11px] text-[#1a1a2e] outline-none"
                    type="number" min="0" max="10" step=".5" value={bleed}
                    onChange={(e) => setBleed(Number(e.target.value))}
                    data-testid="input-print-bleed"
                  />
                </label>
              </div>
            </Panel>

            <Panel>
              <div className="p-4 sm:p-5">
                <div className="pixel-font text-[8px] uppercase tracking-wider text-[#1a1a2e]/60">Export summary</div>
                <div className="mt-2 text-2xl font-bold tracking-tight text-[#1a1a2e]" data-testid="text-export-count">
                  {generatedCards.length}
                </div>
                <p className="mt-1 text-[10px] leading-4 text-[#1a1a2e]/70">
                  Cards in the current browser batch. Generate all to include every valid row.
                </p>
                <button
                  type="button"
                  className="pixel-btn mt-3 inline-flex w-full items-center justify-center gap-2 bg-[#ffe566] px-3 py-2 text-[10px] font-bold text-[#1a1a2e]"
                  onClick={() => void exportPdf()}
                  data-testid="button-export-pdf"
                >
                  <Printer size={13} /> Export {cardsPerPage}-up print PDF
                </button>
              </div>
            </Panel>
          </section>
        )}
      </div>

      <div className="mx-auto mt-5 flex max-w-[1480px] items-center justify-between border-t-[3px] border-[#1a1a2e]/20 pt-3 text-[9px] text-[#1a1a2e]/50">
        <span className="pixel-font uppercase tracking-wider">Local-first rendering · no upload</span>
        <span>SVG · CSV / XLSX · ZIP · PDF</span>
      </div>
    </main>
  );
}

function GalleryCard({
  card,
  selected,
  onClick,
}: {
  card: GeneratedCard;
  selected: boolean;
  onClick: () => void;
}) {
  const [url, setUrl] = useState('');
  useEffect(() => {
    const next = URL.createObjectURL(card.pngBlob);
    setUrl(next);
    return () => URL.revokeObjectURL(next);
  }, [card.pngBlob]);

  return (
    <button
      type="button"
      className={`pixel-border-sm bg-[#fff8e7] p-2 text-left transition-transform hover:-translate-y-0.5 ${
        selected ? 'ring-2 ring-[#ffe566]' : ''
      }`}
      onClick={onClick}
      data-testid={`button-generated-card-${card.rowIndex + 1}`}
    >
      <div className="flex h-24 items-center justify-center bg-[#7ec8f5]/20 p-1">
        <img
          src={url}
          alt={`Generated card row ${card.rowIndex + 1}`}
          className="h-full max-w-full object-contain"
          data-testid={`img-generated-card-${card.rowIndex + 1}`}
        />
      </div>
      <span className="pixel-font mt-1.5 block text-[8px] text-[#1a1a2e]/70">row {card.rowIndex + 1}</span>
    </button>
  );
}

function App() {
  return (
    <Switch>
      <Route path="/" component={BulkIdStudio} />
      <Route>
        <div
          className="grid min-h-[100dvh] place-items-center text-sm text-[#1a1a2e]"
          style={{ backgroundImage: 'linear-gradient(180deg, #7ec8f5 0%, #b8e0f8 40%, #e8f4fc 70%, #f0e6c8 100%)' }}
        >
          <div className="pixel-panel p-6 text-center">
            <div className="text-4xl mb-3">💥</div>
            <div className="pixel-font text-[13px]">PAGE NOT FOUND</div>
          </div>
        </div>
      </Route>
    </Switch>
  );
}

export default App;
