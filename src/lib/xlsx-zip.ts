import JSZip from 'jszip';
import {
  XlsxLoadOptions,
  XlsxPreflightResult,
  XlsxUnsupportedFeature,
  XlsxUnsupportedFeatureType,
} from './xlsx-types';

const DEFAULT_MAX_FILE_SIZE = 50 * 1024 * 1024; // 50 MB
const DEFAULT_MAX_UNCOMPRESSED_SIZE = 200 * 1024 * 1024; // 200 MB
const DEFAULT_MAX_ENTRIES = 10000;

interface FeatureRule {
  type: XlsxUnsupportedFeatureType;
  test: (path: string) => boolean;
  description: string;
}

const FEATURE_RULES: FeatureRule[] = [
  {
    type: 'macro',
    test: (p) =>
      p === 'xl/vbaProject.bin' ||
      p === 'xl/vbaProjectSignature.bin' ||
      p.startsWith('xl/macrosheets/') ||
      p.startsWith('xl/ctrlProps/'),
    description: 'VBA Macros / Automation Scripts',
  },
  {
    type: 'chart',
    test: (p) =>
      p.startsWith('xl/charts/') ||
      p.startsWith('xl/chartsheets/') ||
      p.startsWith('xl/drawings/chart'),
    description: 'Native Excel Charts',
  },
  {
    type: 'pivotTable',
    test: (p) =>
      p.startsWith('xl/pivotTables/') ||
      p.startsWith('xl/pivotCache/'),
    description: 'Pivot Tables & Pivot Caches',
  },
  {
    type: 'externalLink',
    test: (p) => p.startsWith('xl/externalLinks/'),
    description: 'External Workbook Links',
  },
  {
    type: 'embeddedObject',
    test: (p) => p.startsWith('xl/embeddings/'),
    description: 'Embedded OLE Objects / Packages',
  },
  {
    type: 'dataConnection',
    test: (p) =>
      p === 'xl/connections.xml' ||
      p.startsWith('xl/queryTables/'),
    description: 'Data Connections / Power Query',
  },
  {
    type: 'customXml',
    test: (p) => p.startsWith('customXml/'),
    description: 'Custom XML Schemas / Data Parts',
  },
  {
    type: 'slicer',
    test: (p) =>
      p.startsWith('xl/slicers/') ||
      p.startsWith('xl/slicerCaches/'),
    description: 'Interactive Slicers',
  },
  {
    type: 'modelExtension',
    test: (p) =>
      p.startsWith('xl/model/') ||
      p.startsWith('xl/richData/') ||
      p === 'xl/metadata.xml',
    description: 'Data Model / Rich Data Extensions',
  },
];

/**
 * Inspects a zip buffer representing an XLSX file before loading into ExcelJS.
 * Detects features that ExcelJS is known to strip/corrupt.
 */
export async function inspectZipBuffer(
  buffer: Buffer,
  options?: XlsxLoadOptions
): Promise<XlsxPreflightResult> {
  const maxFileSize = options?.maxFileSize ?? DEFAULT_MAX_FILE_SIZE;
  const maxUncompressed = options?.maxUncompressedSize ?? DEFAULT_MAX_UNCOMPRESSED_SIZE;
  const maxEntries = options?.maxEntries ?? DEFAULT_MAX_ENTRIES;

  if (buffer.length > maxFileSize) {
    throw new Error(
      `XLSX file size (${buffer.length} bytes) exceeds maximum allowed size of ${maxFileSize} bytes.`
    );
  }

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer, { checkCRC32: true });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(`Failed to parse XLSX as a valid ZIP archive: ${msg}`);
  }

  const entries = Object.keys(zip.files);
  if (entries.length > maxEntries) {
    throw new Error(
      `XLSX ZIP archive contains ${entries.length} entries, exceeding limit of ${maxEntries}.`
    );
  }

  const unsupportedMap = new Map<string, XlsxUnsupportedFeature>();
  let totalUncompressedSize = 0;
  const warnings: string[] = [];

  for (const entryName of entries) {
    const file = zip.files[entryName];
    if (!file) continue;

    // Normalize path separators to POSIX
    const normalizedPath = entryName.replace(/\\/g, '/');

    // Track uncompressed size for safety against zip bombs (fail closed if invalid)
    if (!file.dir) {
      const rawData = (file as unknown as { _data?: { uncompressedSize?: unknown } })._data;
      const uncompressed = rawData?.uncompressedSize;
      if (typeof uncompressed !== 'number' || !Number.isFinite(uncompressed) || uncompressed < 0) {
        throw new Error(
          `Cannot determine uncompressed size for ZIP entry "${normalizedPath}". Preflight inspection fails closed for safety.`
        );
      }
      totalUncompressedSize += uncompressed;

      if (totalUncompressedSize > maxUncompressed) {
        throw new Error(
          `XLSX ZIP total uncompressed size exceeds limit of ${maxUncompressed} bytes (potential zip bomb).`
        );
      }
    }

    // Check against unsupported feature rules
    for (const rule of FEATURE_RULES) {
      if (rule.test(normalizedPath)) {
        if (!unsupportedMap.has(normalizedPath)) {
          unsupportedMap.set(normalizedPath, {
            type: rule.type,
            description: `${rule.description} (${normalizedPath})`,
            path: normalizedPath,
          });
        }
      }
    }
  }

  // Extract sheet names from xl/workbook.xml if available
  const sheetNames: string[] = [];
  const workbookXmlFile = zip.file('xl/workbook.xml');
  if (workbookXmlFile) {
    try {
      const xmlContent = await workbookXmlFile.async('text');
      // Match <sheet ... name="SheetName" .../>
      const sheetTagRegex = /<sheet\b[^>]*\bname="([^"]+)"[^>]*\/?>/gi;
      let match: RegExpExecArray | null;
      while ((match = sheetTagRegex.exec(xmlContent)) !== null) {
        if (match[1]) {
          // Decode basic XML entities in sheet name
          const decoded = match[1]
            .replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'");
          sheetNames.push(decoded);
        }
      }
    } catch {
      warnings.push('Could not parse sheet names from xl/workbook.xml during preflight.');
    }
  }

  const unsupportedFeatures = Array.from(unsupportedMap.values());
  if (unsupportedFeatures.length > 0) {
    warnings.push(
      `Detected ${unsupportedFeatures.length} unsupported feature item(s) in workbook: ` +
        unsupportedFeatures.map((f) => f.description).join('; ')
    );
  }

  return {
    hasUnsupportedFeatures: unsupportedFeatures.length > 0,
    unsupportedFeatures,
    sheetNames,
    entryCount: entries.length,
    totalUncompressedSize,
    warnings,
  };
}
