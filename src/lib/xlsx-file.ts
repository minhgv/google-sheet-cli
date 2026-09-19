import * as crypto from 'crypto';
import * as fs from 'fs';
import * as path from 'path';
import { XlsxSaveOptions, XlsxSaveResult } from './xlsx-types';

/**
 * Computes SHA-256 hash of a buffer.
 */
export function computeSha256(buffer: Buffer): string {
  return crypto.createHash('sha256').update(buffer).digest('hex');
}

/**
 * Computes SHA-256 hash of a file on disk.
 */
export async function computeFileSha256(filePath: string): Promise<string> {
  const buffer = await fs.promises.readFile(filePath);
  return computeSha256(buffer);
}

/**
 * Determines if two paths point to the exact same file (via realpath and inode/device check).
 */
export async function isSameFile(pathA: string, pathB: string): Promise<boolean> {
  try {
    const [realA, realB] = await Promise.all([
      fs.promises.realpath(pathA).catch(() => null),
      fs.promises.realpath(pathB).catch(() => null),
    ]);

    if (realA && realB && realA === realB) {
      return true;
    }

    const [statA, statB] = await Promise.all([
      fs.promises.stat(pathA).catch(() => null),
      fs.promises.stat(pathB).catch(() => null),
    ]);

    if (statA && statB) {
      return statA.dev === statB.dev && statA.ino === statB.ino;
    }

    return false;
  } catch {
    return false;
  }
}

/**
 * Creates a backup copy of a file before overwriting.
 */
export async function createBackupCopy(
  sourcePath: string,
  customBackupPath?: string
): Promise<string> {
  const destPath = customBackupPath || `${sourcePath}.bak`;
  await fs.promises.copyFile(sourcePath, destPath);
  return destPath;
}

/**
 * Saves a buffer to a target path atomically with overwrite/in-place safeguards,
 * conflict verification, and automatic cleanup of temporary files on error.
 */
export async function saveBufferAtomic(
  buffer: Buffer,
  targetPath: string,
  options?: XlsxSaveOptions,
  context?: { originalPath?: string; originalHash?: string }
): Promise<XlsxSaveResult> {
  if (!targetPath || typeof targetPath !== 'string' || !targetPath.trim()) {
    throw new Error('Target save path must be a non-empty string.');
  }

  const resolvedTarget = path.resolve(targetPath);
  const targetDir = path.dirname(resolvedTarget);

  // Ensure target directory exists
  await fs.promises.mkdir(targetDir, { recursive: true });

  const targetExists = await fs.promises
    .access(resolvedTarget, fs.constants.F_OK)
    .then(() => true)
    .catch(() => false);

  let isSameAsOriginal = false;
  if (context?.originalPath) {
    isSameAsOriginal = await isSameFile(resolvedTarget, path.resolve(context.originalPath));
  }

  // Safety check 1: Overwrite protection
  if (targetExists) {
    const isExplicitInPlace = options?.inPlace === true;
    const isExplicitOverwrite = options?.overwrite === true;

    if (!isExplicitInPlace && !isExplicitOverwrite) {
      throw new Error(
        `Target file "${resolvedTarget}" already exists. Pass overwrite: true or inPlace: true to replace.`
      );
    }

    // Safety check 2: Conflict detection via SHA-256 fingerprint
    const expectedHash = options?.expectedHash || (isSameAsOriginal ? context?.originalHash : undefined);
    if (expectedHash) {
      const currentDiskHash = await computeFileSha256(resolvedTarget);
      if (currentDiskHash !== expectedHash) {
        throw new Error(
          `Conflict detected: destination file has changed on disk since it was loaded. Expected hash ${expectedHash}, found ${currentDiskHash}.`
        );
      }
    }
  }

  // Backup creation (required when inPlace is true, or when options.backup is true)
  let backupCreated: string | undefined;
  if (targetExists && (options?.backup === true || options?.inPlace === true)) {
    backupCreated = await createBackupCopy(resolvedTarget, options?.backupPath);
  }

  // Atomic write via temporary file in the same directory
  const randomSuffix = crypto.randomBytes(6).toString('hex');
  const tempFileName = `.${path.basename(resolvedTarget)}.${process.pid}.${Date.now()}.${randomSuffix}.tmp`;
  const tempFilePath = path.join(targetDir, tempFileName);

  try {
    // Write to temporary file
    await fs.promises.writeFile(tempFilePath, buffer, { mode: 0o644 });

    // Atomically replace target
    await fs.promises.rename(tempFilePath, resolvedTarget);
  } catch (err: unknown) {
    // Cleanup temporary file on failure
    await fs.promises.unlink(tempFilePath).catch(() => {});
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(`Failed to safely save XLSX to "${resolvedTarget}": ${msg}`);
  }

  const sha256Hash = computeSha256(buffer);

  return {
    savedPath: resolvedTarget,
    bytesWritten: buffer.length,
    sha256Hash,
    backupCreated,
    isOverwritten: targetExists,
  };
}
