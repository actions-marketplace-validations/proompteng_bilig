import { closeSync, openSync, readSync, writeSync } from 'node:fs'

import {
  assertZip32Size,
  centralDirectoryHeader,
  copyCompressedSourceEntryToZipOutput,
  crc32,
  deflateZipEntry,
  encodeZipPath,
  endOfCentralDirectory,
  localFileHeader,
  readUnmaterializedSourceEntry,
  zipDeflateCompressionMethod,
  type FilePreparedZipEntry,
  type SourcePreservingZip,
  type SourcePreservingZipEntryRecord,
  type SourcePreservingZipOptions,
} from './source-preserving-zip.js'

function currentZipDosTimeParts(): { readonly time: number; readonly date: number } {
  const now = new Date()
  const year = Math.max(1980, Math.min(2099, now.getFullYear()))
  return {
    time: (now.getHours() << 11) | (now.getMinutes() << 5) | Math.trunc(now.getSeconds() / 2),
    date: ((year - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate(),
  }
}

export function writeAllSync(fd: number, chunk: Uint8Array): void {
  let offset = 0
  while (offset < chunk.byteLength) {
    offset += writeSync(fd, chunk, offset, chunk.byteLength - offset)
  }
}

export function zipSourcePreservingEntriesToFile(
  zip: SourcePreservingZip,
  preparedEntries: ReadonlyMap<string, FilePreparedZipEntry>,
  outputPath: string,
  options: SourcePreservingZipOptions = {},
): number {
  let outputByteLength = 0
  const records: SourcePreservingZipEntryRecord[] = []
  const { time, date } = options.dosTime ?? currentZipDosTimeParts()
  const paths = [...new Set([...Object.keys(zip), ...preparedEntries.keys()])]
  const fd = openSync(outputPath, 'w')
  const pushOutput = (chunk: Uint8Array): void => {
    writeAllSync(fd, chunk)
    outputByteLength += chunk.byteLength
  }
  try {
    for (const path of paths) {
      const preparedEntry = preparedEntries.get(path)
      const sourceEntry = preparedEntry ? null : readUnmaterializedSourceEntry(zip, path)
      const bytes = preparedEntry || sourceEntry ? undefined : zip[path]
      if (!preparedEntry && !sourceEntry && !bytes) {
        continue
      }
      assertZip32Size(outputByteLength)
      const { bytes: pathBytes, utf8 } = encodeZipPath(path)
      const localHeaderOffset = outputByteLength
      if (preparedEntry) {
        assertZip32Size(preparedEntry.uncompressedSize)
        assertZip32Size(preparedEntry.compressedSize)
        const record: SourcePreservingZipEntryRecord = {
          pathBytes,
          utf8,
          flags: 0,
          compressionMethod: zipDeflateCompressionMethod,
          crc: preparedEntry.crc,
          compressedSize: preparedEntry.compressedSize,
          uncompressedSize: preparedEntry.uncompressedSize,
          localHeaderOffset,
        }
        pushOutput(localFileHeader(record, time, date))
        copyFileToZipOutput(preparedEntry.compressedPath, pushOutput)
        records.push(record)
      } else if (sourceEntry) {
        assertZip32Size(sourceEntry.uncompressedSize)
        assertZip32Size(sourceEntry.compressedSize)
        const record: SourcePreservingZipEntryRecord = {
          pathBytes,
          utf8,
          flags: 0,
          compressionMethod: sourceEntry.compressionMethod,
          crc: sourceEntry.crc,
          compressedSize: sourceEntry.compressedSize,
          uncompressedSize: sourceEntry.uncompressedSize,
          localHeaderOffset,
        }
        pushOutput(localFileHeader(record, time, date))
        copyCompressedSourceEntryToZipOutput(sourceEntry, pushOutput)
        records.push(record)
      } else {
        const uncompressedSize = bytes!.byteLength
        const compressed = deflateZipEntry(bytes!)
        const compressedSize = compressed.byteLength
        const crc = crc32(bytes!)
        assertZip32Size(uncompressedSize)
        assertZip32Size(compressedSize)
        const record: SourcePreservingZipEntryRecord = {
          pathBytes,
          utf8,
          flags: 0,
          compressionMethod: zipDeflateCompressionMethod,
          crc,
          compressedSize,
          uncompressedSize,
          localHeaderOffset,
        }
        pushOutput(localFileHeader(record, time, date))
        pushOutput(compressed)
        records.push(record)
      }
      delete zip[path]
    }

    const centralDirectoryOffset = outputByteLength
    for (const record of records) {
      pushOutput(centralDirectoryHeader(record, time, date))
    }
    const centralDirectorySize = outputByteLength - centralDirectoryOffset
    pushOutput(endOfCentralDirectory(records.length, centralDirectorySize, centralDirectoryOffset))
  } finally {
    closeSync(fd)
  }
  return outputByteLength
}

function copyFileToZipOutput(path: string, pushOutput: (chunk: Uint8Array) => void): void {
  const fd = openSync(path, 'r')
  const scratch = new Uint8Array(64 * 1024)
  try {
    let bytesRead = 0
    do {
      bytesRead = readSync(fd, scratch, 0, scratch.byteLength, null)
      if (bytesRead > 0) {
        pushOutput(scratch.subarray(0, bytesRead))
      }
    } while (bytesRead > 0)
  } finally {
    closeSync(fd)
  }
}
