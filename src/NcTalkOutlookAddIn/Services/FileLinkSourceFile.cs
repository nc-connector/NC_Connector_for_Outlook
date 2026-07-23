// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Applies the planned source snapshot to every local file read.
    internal static class FileLinkSourceFile
    {
        internal const int BufferSize = 81920;

        internal static void ValidateSnapshot(FileLinkPlannedFile file)
        {
            if (file == null)
            {
                ThrowSourceChanged();
            }

            var current = new FileInfo(file.LocalPath);
            current.Refresh();
            if (!current.Exists
                || current.Length != file.Length
                || current.LastWriteTimeUtc != file.LastWriteTimeUtc)
            {
                ThrowSourceChanged();
            }
        }

        internal static string ComputeMd5Hex(FileLinkPlannedFile file)
        {
            ValidateSnapshot(file);
            byte[] checksum;
            using (var md5 = MD5.Create())
            using (FileStream source = OpenRead(file))
            {
                checksum = md5.ComputeHash(source);
            }
            ValidateSnapshot(file);
            return BitConverter
                .ToString(checksum)
                .Replace("-", string.Empty)
                .ToLowerInvariant();
        }

        internal static void WriteRange(
            FileLinkPlannedFile file,
            Stream destination,
            long offset,
            long length,
            byte[] buffer,
            CancellationToken cancellationToken,
            Action<long> reportBytes)
        {
            if (destination == null)
            {
                throw new ArgumentNullException("destination");
            }
            if (buffer == null || buffer.Length == 0)
            {
                throw new ArgumentException(
                    "A non-empty transfer buffer is required.",
                    "buffer");
            }
            if (offset < 0
                || length < 0
                || file == null
                || offset > file.Length
                || length > file.Length - offset)
            {
                throw new ArgumentOutOfRangeException("length");
            }

            ValidateSnapshot(file);
            long transferred = 0;
            using (FileStream source = OpenRead(file))
            {
                source.Seek(offset, SeekOrigin.Begin);
                long remaining = length;
                while (remaining > 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int toRead = (int)Math.Min(
                        buffer.Length,
                        remaining);
                    int bytesRead = source.Read(buffer, 0, toRead);
                    if (bytesRead <= 0)
                    {
                        throw new EndOfStreamException(
                            "Unexpected end of file during upload.");
                    }

                    destination.Write(buffer, 0, bytesRead);
                    remaining -= bytesRead;
                    transferred += bytesRead;
                    if (reportBytes != null)
                    {
                        reportBytes(transferred);
                    }
                }
            }
            ValidateSnapshot(file);
        }

        private static FileStream OpenRead(FileLinkPlannedFile file)
        {
            return new FileStream(
                file.LocalPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                BufferSize,
                FileOptions.SequentialScan);
        }

        private static void ThrowSourceChanged()
        {
            throw new TalkServiceException(
                Strings.FileLinkUploadSourceChanged,
                false,
                0,
                null);
        }
    }
}
