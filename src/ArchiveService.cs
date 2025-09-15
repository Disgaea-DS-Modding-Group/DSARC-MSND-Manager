using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
namespace DisgaeaDS_Manager
{
    internal interface IArchiveService
    {
        Task<Collection<Entry>> LoadArchiveAsync(string archivePath, CancellationToken ct = default);
        Task SaveArchiveAsync(string path, ArchiveType type, IList<Entry> entries, string srcFolder, IProgress<(int current, int total)> progress, CancellationToken ct);
        Task<ImportResult> InspectFolderForImportAsync(string folder, CancellationToken ct = default);
        Task<(string outBase, List<string> mapperLines)> ExtractAllAsync(string archivePath, ArchiveType filetype, string destFolder, IProgress<(int current, int total)>? progress = null, CancellationToken ct = default);
        Task ExtractItemAsync(string archivePath, ArchiveType filetype, Entry entry, string destFolder, CancellationToken ct = default);
        Task ExtractChunkItemAsync(string archivePath, Entry parentEntry, Entry chunkEntry, string destFolder, CancellationToken ct = default);
        Task<byte[]> ReplaceChunkItemAsync(string archivePath, Entry parentEntry, Entry chunkEntry, string replacementFilePath, string srcFolder, CancellationToken ct = default);
        Task<byte[]> RebuildNestedFromFolderAsync(string folder, CancellationToken ct = default);
        Task NestedExtractBufferAsync(byte[] buf, string outdir, string baseLabel, IProgress<(int current, int total)>? progress = null, CancellationToken ct = default);
        Task<Collection<Entry>> ParseDsarcFromBufferAsync(byte[] buf, CancellationToken ct = default);
        Task<byte[]> BuildDsarcFromFolderAsync(string folder, CancellationToken ct = default);
        Task<byte[]> BuildMsndFromFolderAsync(string folder, CancellationToken ct = default);
        Task WriteFileAsync(string path, byte[] data, CancellationToken ct = default);
        Task<byte[]> ReadFileAsync(string path, CancellationToken ct = default);
        Task<byte[]> ReadRangeAsync(string path, long offset, int size, CancellationToken ct = default);
        Task CopyFileToFolderAsync(string sourceFilePath, string destFolder, CancellationToken ct = default);
    }
    internal class ArchiveService : IArchiveService
    {
        private readonly TupleComparer _tupleComparer = new();
        private static readonly char[] separator = new[] { '=' };
        internal static readonly char[] separatorArray = new[] { '=' };
        internal static readonly string[] sourceArray = new[] { ".sseq", ".sbnk", ".swar" };

        public async Task<Collection<Entry>> LoadArchiveAsync(string archivePath, CancellationToken ct = default)
        {
            return await Task.Run(() =>
            {
                ArchiveType type = Detector.FromFile(archivePath);
                return type == ArchiveType.MSND
                    ? Msnd.Parse(File.ReadAllBytes(archivePath), Path.GetFileNameWithoutExtension(archivePath))
                    : Dsarc.Parse(archivePath);
            }, ct).ConfigureAwait(false);
        }
        public async Task SaveArchiveAsync(string path, ArchiveType type, IList<Entry> entries, string srcFolder, IProgress<(int current, int total)>? progress = null, CancellationToken ct = default)
        {
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentNullException(nameof(path));
            }
            if (entries == null || entries.Count == 0)
            {
                throw new ArgumentException("No entries to save", nameof(entries));
            }
            if (string.IsNullOrEmpty(srcFolder))
            {
                throw new ArgumentNullException(nameof(srcFolder));
            }
            await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                if (type == ArchiveType.MSND)
                {
                    Dictionary<string, byte[]> chunks = new(StringComparer.OrdinalIgnoreCase);
                    int total = Msnd.MSNDORDER.Length;
                    for (int i = 0; i < total; i++)
                    {
                        ct.ThrowIfCancellationRequested();
                        string ext = Msnd.MSNDORDER[i];
                        string? sourcePath = entries.FirstOrDefault(e => string.Equals(e.Path.Extension, ext, StringComparison.OrdinalIgnoreCase))?.Path.ToString();
                        sourcePath = sourcePath != null ? Path.Combine(srcFolder, sourcePath) : Directory.GetFiles(srcFolder, $"*{ext}", SearchOption.TopDirectoryOnly).FirstOrDefault();
                        if (sourcePath == null)
                        {
                            throw new FileNotFoundException($"Missing {ext} in {srcFolder}");
                        }
                        chunks[ext] = File.ReadAllBytes(sourcePath);
                        progress?.Report((i + 1, total));
                    }
                    File.WriteAllBytes(path, Msnd.Build(chunks));
                }
                else
                {
                    string mappingPath = Path.Combine(srcFolder, "mapper.txt");
                    if (File.Exists(mappingPath))
                    {
                        List<string> missing = [];
                        Collection<Tuple<string, byte[]>> pairs = [];
                        string[] lines = File.ReadAllLines(mappingPath, Encoding.UTF8);
                        int total = lines.Length;
                        for (int i = 0; i < lines.Length; i++)
                        {
                            ct.ThrowIfCancellationRequested();
                            string ln = lines[i];
                            if (!ln.Contains('=', StringComparison.Ordinal))
                            {
                                progress?.Report((i + 1, total));
                                continue;
                            }
                            string[] parts = ln.Split(separator, 2);
                            string left = parts[0].Trim();
                            string right = parts[1].Trim();
                            string candidate = Path.Combine(srcFolder, right);
                            if (Directory.Exists(candidate))
                            {
                                byte[] childBuf = RebuildNestedFromFolderAsync(candidate, ct).GetAwaiter().GetResult();
                                if (childBuf == null)
                                {
                                    missing.Add(right);
                                    progress?.Report((i + 1, total));
                                    continue;
                                }
                                pairs.Add(Tuple.Create(left, childBuf));
                                progress?.Report((i + 1, total));
                                continue;
                            }
                            if (File.Exists(candidate))
                            {
                                pairs.Add(Tuple.Create(left, File.ReadAllBytes(candidate)));
                                progress?.Report((i + 1, total));
                                continue;
                            }
                            string[] matches = Directory.GetFiles(srcFolder, Path.GetFileName(right), SearchOption.AllDirectories);
                            if (matches?.Length > 0)
                            {
                                pairs.Add(Tuple.Create(left, File.ReadAllBytes(matches[0])));
                                progress?.Report((i + 1, total));
                                continue;
                            }
                            missing.Add(right);
                            progress?.Report((i + 1, total));
                        }
                        if (missing.Count > 0)
                        {
                            throw new FileNotFoundException($"The mapping file references files or folders that are missing:\n{string.Join("\n ", missing)}");
                        }
                        File.WriteAllBytes(path, Dsarc.BuildFromPairs(pairs));
                    }
                    else
                    {
                        Collection<Tuple<string, byte[]>> pairs = [];
                        int total = entries.Count;
                        for (int i = 0; i < entries.Count; i++)
                        {
                            ct.ThrowIfCancellationRequested();
                            Entry e = entries[i];
                            string candidate = Path.Combine(srcFolder, e.Path.ToString());
                            if (Directory.Exists(candidate))
                            {
                                byte[] childBuf = RebuildNestedFromFolderAsync(candidate, ct).GetAwaiter().GetResult();
                                if (childBuf == null)
                                {
                                    throw new InvalidOperationException($"Failed to rebuild nested archive at {candidate}");
                                }
                                pairs.Add(Tuple.Create(e.Path.ToString(), childBuf));
                                progress?.Report((i + 1, total));
                                continue;
                            }
                            if (File.Exists(candidate))
                            {
                                pairs.Add(Tuple.Create(e.Path.ToString(), File.ReadAllBytes(candidate)));
                                progress?.Report((i + 1, total));
                                continue;
                            }
                            string[] matches = Directory.GetFiles(srcFolder, Path.GetFileName(e.Path.ToString()), SearchOption.AllDirectories);
                            if (matches?.Length > 0)
                            {
                                pairs.Add(Tuple.Create(e.Path.ToString(), File.ReadAllBytes(matches[0])));
                                progress?.Report((i + 1, total));
                                continue;
                            }
                            throw new FileNotFoundException($"Missing source file or folder for entry: {e.Path} (expected at {candidate})");
                        }
                        File.WriteAllBytes(path, Dsarc.BuildFromPairs(pairs));
                        progress?.Report((1, 1));
                    }
                }
            }, ct).ConfigureAwait(false);
        }
        public async Task<ImportResult> InspectFolderForImportAsync(string folder, CancellationToken ct = default)
        {
            return await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                if (!Directory.Exists(folder))
                {
                    throw new DirectoryNotFoundException(folder);
                }
                ImportResult result = new() { SourceFolder = folder };
                string mappingFile = Path.Combine(folder, "mapper.txt");
                if (File.Exists(mappingFile))
                {
                    foreach (string ln in File.ReadAllLines(mappingFile, Encoding.UTF8))
                    {
                        ct.ThrowIfCancellationRequested();
                        if (!ln.Contains('=', StringComparison.Ordinal))
                        {
                            continue;
                        }
                        string[] parts = ln.Split(separator, 2);
                        string left = parts[0].Trim();
                        string right = parts[1].Trim();
                        string candidate = Path.Combine(folder, right);
                        if (Directory.Exists(candidate))
                        {
                            Entry entry = new(new FileInfo(left));
                            try
                            {
                                byte[] buf = RebuildNestedFromFolderAsync(candidate, ct).GetAwaiter().GetResult();
                                if (buf != null && buf.Length >= 4 && buf.Take(4).SequenceEqual(Msnd.MAGICMSND))
                                {
                                    entry.IsMsnd = true;
                                    foreach (Entry child in Msnd.Parse(buf, Path.GetFileNameWithoutExtension(entry.Path.Name)))
                                    {
                                        entry.Children.Add(child);
                                    }
                                }
                            }
                            catch (IOException) { }
                            catch (UnauthorizedAccessException) { }
                            result.Entries.Add(entry);
                            continue;
                        }
                        if (File.Exists(candidate))
                        {
                            result.Entries.Add(new Entry(new FileInfo(left), (int)new FileInfo(candidate).Length, 0));
                            continue;
                        }
                        string[] matches = Directory.GetFiles(folder, Path.GetFileName(right), SearchOption.AllDirectories);
                        if (matches?.Length > 0)
                        {
                            result.Entries.Add(new Entry(new FileInfo(left), (int)new FileInfo(matches[0]).Length, 0));
                        }
                        else
                        {
                            result.Entries.Add(new Entry(new FileInfo(left)));
                        }
                    }
                    result.FileType = ArchiveType.DSARC;
                    return result;
                }
                List<string> topLevel = Directory.GetFileSystemEntries(folder, "*", SearchOption.TopDirectoryOnly)
                    .Where(p => !string.Equals(Path.GetFileName(p), "mapper.txt", StringComparison.OrdinalIgnoreCase)).ToList();
                List<string> allFiles = Directory.GetFiles(folder, "*", SearchOption.AllDirectories)
                    .Where(f => !f.EndsWith("mapper.txt", StringComparison.OrdinalIgnoreCase)).OrderBy(f => f).ToList();
                HashSet<string> stems = [.. allFiles.Select(f => Path.GetFileNameWithoutExtension(f))];
                string? stemCandidate = stems.Count == 1 ? stems.First() : null;
                HashSet<string> exts = [.. allFiles.Select(f => Path.GetExtension(f).ToUpperInvariant())];
                HashSet<string> msndOrderUpper = [.. Msnd.MSNDORDER.Select(x => x.ToUpperInvariant())];
                bool onlyThree = exts.SetEquals(msndOrderUpper) && allFiles.Count(f => msndOrderUpper.Contains(Path.GetExtension(f).ToUpperInvariant())) == 3;
                if ((stemCandidate != null && Msnd.MSNDORDER.All(x => exts.Contains(x.ToUpperInvariant()))) || onlyThree)
                {
                    result.FileType = ArchiveType.MSND;
                    foreach (string ext in Msnd.MSNDORDER)
                    {
                        string? chosen = null;
                        if (stemCandidate != null)
                        {
                            chosen = allFiles.FirstOrDefault(f => Path.GetFileNameWithoutExtension(f) == stemCandidate && Path.GetExtension(f).Equals(ext, StringComparison.OrdinalIgnoreCase));
                        }
                        chosen ??= allFiles.FirstOrDefault(f => Path.GetExtension(f).Equals(ext, StringComparison.OrdinalIgnoreCase));
                        if (chosen == null)
                        {
                            throw new FileNotFoundException($"Missing expected MSND file {ext}");
                        }
                        result.Entries.Add(new Entry(new FileInfo(Path.GetFileName(chosen))));
                    }
                    return result;
                }
                result.FileType = ArchiveType.DSARC;
                foreach (string p in topLevel)
                {
                    ct.ThrowIfCancellationRequested();
                    if (Directory.Exists(p))
                    {
                        Entry entry = new(new FileInfo(Path.GetFileName(p)));
                        try
                        {
                            byte[] buf = RebuildNestedFromFolderAsync(p, ct).GetAwaiter().GetResult();
                            if (buf != null && buf.Length >= 4 && buf.Take(4).SequenceEqual(Msnd.MAGICMSND))
                            {
                                entry.IsMsnd = true;
                                foreach (Entry child in Msnd.Parse(buf, Path.GetFileNameWithoutExtension(entry.Path.Name)))
                                {
                                    entry.Children.Add(child);
                                }
                            }
                        }
                        catch (IOException) { }
                        catch (UnauthorizedAccessException) { }
                        result.Entries.Add(entry);
                    }
                    else if (File.Exists(p))
                    {
                        string rel = GetRelativePath(folder, p);
                        result.Entries.Add(new Entry(new FileInfo(rel)));
                    }
                }
                return result;
            }, ct).ConfigureAwait(false);
        }
        public async Task<(string outBase, List<string> mapperLines)> ExtractAllAsync(string archivePath, ArchiveType filetype, string destFolder, IProgress<(int current, int total)>? progress = null, CancellationToken ct = default)
        {
            return string.IsNullOrEmpty(archivePath)
                ? throw new ArgumentNullException(nameof(archivePath))
                : string.IsNullOrEmpty(destFolder)
                ? throw new ArgumentNullException(nameof(destFolder))
                : ((string outBase, List<string> mapperLines))await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                Collection<Entry> entries = filetype == ArchiveType.MSND ? Msnd.Parse(File.ReadAllBytes(archivePath), Path.GetFileNameWithoutExtension(archivePath)) : Dsarc.Parse(archivePath);
                string baseOut = Path.Combine(destFolder, Path.GetFileNameWithoutExtension(archivePath));
                _ = Directory.CreateDirectory(baseOut);
                if (filetype == ArchiveType.MSND)
                {
                    byte[] buf = File.ReadAllBytes(archivePath);
                    string baseName = Path.GetFileNameWithoutExtension(archivePath);
                    if (buf.Length >= Msnd.HDRMSND)
                    {
                        byte[] txt = new byte[4];
                        Array.Copy(buf, 44, txt, 0, Math.Min(4, buf.Length - 44));
                        File.WriteAllBytes(Path.Combine(baseOut, baseName + ".txt"), txt);
                    }
                    int total = entries.Count;
                    for (int i = 0; i < entries.Count; i++)
                    {
                        ct.ThrowIfCancellationRequested();
                        Entry e = entries[i];
                        byte[] data = new byte[e.Size];
                        Array.Copy(buf, e.Offset, data, 0, e.Size);
                        string target = Path.Combine(baseOut, e.Path.Name);
                        _ = Directory.CreateDirectory(Path.GetDirectoryName(target) ?? baseOut);
                        File.WriteAllBytes(target, data);
                        progress?.Report((i + 1, total));
                    }
                    return (baseOut, new List<string>());
                }
                else
                {
                    long archiveSize = new FileInfo(archivePath).Length;
                    long dataStart = Dsarc.HDRDSARC + (entries.Count * (Helpers.NAMESZ + Dsarc.ENTRYINFOSZ));
                    var ranges = entries.Select((e, idx) => new
                    {
                        Entry = e,
                        Index = idx,
                        Start = (long)e.Offset,
                        End = (long)e.Offset + e.Size
                    }).Where(x => x.Start >= dataStart && x.End <= archiveSize && x.Entry.Size >= 0 && x.End >= x.Start)
                      .OrderBy(x => x.Start)
                      .ToList();
                    int total = ranges.Count;
                    string[] extractedNames = new string[entries.Count];
                    using (FileStream fs = new(archivePath, FileMode.Open, FileAccess.Read))
                    {
                        for (int i = 0; i < ranges.Count; i++)
                        {
                            ct.ThrowIfCancellationRequested();
                            var range = ranges[i];
                            _ = fs.Seek(range.Start, SeekOrigin.Begin);
                            byte[] data = new byte[range.End - range.Start];
                            _ = fs.Read(data, 0, data.Length);
                            string ext = Helpers.GuessExtByMagic(data, range.Entry.Path.Extension);
                            string baseName = Path.GetFileNameWithoutExtension(range.Entry.Path.Name);
                            string finalName = UniqueOutName(baseName, ext, baseOut, new Dictionary<Tuple<string, string>, int>(_tupleComparer));
                            string target = Path.Combine(baseOut, finalName);
                            _ = Directory.CreateDirectory(Path.GetDirectoryName(target) ?? baseOut);
                            File.WriteAllBytes(target, data);
                            extractedNames[range.Index] = finalName;
                            progress?.Report((i + 1, total));
                        }
                    }
                    List<string> mappingLines = entries.Select((e, idx) => $"{e.Path.Name}={extractedNames[idx] ?? e.Path.Name}").ToList();
                    File.WriteAllText(Path.Combine(baseOut, "mapper.txt"), string.Join("\n", mappingLines), Encoding.UTF8);
                    return (baseOut, mappingLines);
                }
            }, ct).ConfigureAwait(false);
        }
        public async Task ExtractItemAsync(string archivePath, ArchiveType filetype, Entry entry, string destFolder, CancellationToken ct = default)
        {
            if (string.IsNullOrEmpty(archivePath))
            {
                throw new ArgumentNullException(nameof(archivePath));
            }
            ArgumentNullException.ThrowIfNull(entry);
            if (string.IsNullOrEmpty(destFolder))
            {
                throw new ArgumentNullException(nameof(destFolder));
            }
            await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                if (filetype == ArchiveType.MSND)
                {
                    byte[] buf = File.ReadAllBytes(archivePath);
                    byte[] data = new byte[entry.Size];
                    Array.Copy(buf, entry.Offset, data, 0, entry.Size);
                    string target = Path.Combine(destFolder, entry.Path.Name);
                    _ = Directory.CreateDirectory(Path.GetDirectoryName(target) ?? destFolder);
                    File.WriteAllBytes(target, data);
                }
                else
                {
                    using FileStream fs = new(archivePath, FileMode.Open, FileAccess.Read);
                    _ = fs.Seek(entry.Offset, SeekOrigin.Begin);
                    byte[] data = new byte[entry.Size];
                    _ = fs.Read(data, 0, data.Length);
                    string ext = Helpers.GuessExtByMagic(data, Path.GetExtension(entry.Path.Name));
                    string target = Path.Combine(destFolder, Path.GetFileNameWithoutExtension(entry.Path.Name) + ext);
                    _ = Directory.CreateDirectory(Path.GetDirectoryName(target) ?? destFolder);
                    File.WriteAllBytes(target, data);
                }
            }, ct).ConfigureAwait(false);
        }
        public async Task ExtractChunkItemAsync(string archivePath, Entry parentEntry, Entry chunkEntry, string destFolder, CancellationToken ct = default)
        {
            if (parentEntry == null || chunkEntry == null)
            {
                throw new ArgumentNullException(parentEntry == null ? nameof(parentEntry) : nameof(chunkEntry));
            }
            if (string.IsNullOrEmpty(archivePath))
            {
                throw new ArgumentNullException(nameof(archivePath));
            }
            if (string.IsNullOrEmpty(destFolder))
            {
                throw new ArgumentNullException(nameof(destFolder));
            }
            await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                byte[] msndBuf;
                using (FileStream fs = new(archivePath, FileMode.Open, FileAccess.Read))
                {
                    _ = fs.Seek(parentEntry.Offset, SeekOrigin.Begin);
                    msndBuf = new byte[parentEntry.Size];
                    _ = fs.Read(msndBuf, 0, msndBuf.Length);
                }
                byte[] data = new byte[chunkEntry.Size];
                Array.Copy(msndBuf, chunkEntry.Offset, data, 0, chunkEntry.Size);
                string target = Path.Combine(destFolder, chunkEntry.Path.Name);
                _ = Directory.CreateDirectory(Path.GetDirectoryName(target) ?? destFolder);
                File.WriteAllBytes(target, data);
            }, ct).ConfigureAwait(false);
        }
        public async Task<byte[]> ReplaceChunkItemAsync(string archivePath, Entry parentEntry, Entry chunkEntry, string replacementFilePath, string srcFolder, CancellationToken ct = default)
        {
            return parentEntry == null || chunkEntry == null
                ? throw new ArgumentNullException(parentEntry == null ? nameof(parentEntry) : nameof(chunkEntry))
                : string.IsNullOrEmpty(archivePath)
                ? throw new ArgumentNullException(nameof(archivePath))
                : string.IsNullOrEmpty(replacementFilePath)
                ? throw new ArgumentNullException(nameof(replacementFilePath))
                : await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                byte[] newData = File.ReadAllBytes(replacementFilePath);
                byte[] msndBuf;
                using (FileStream fs = new(archivePath, FileMode.Open, FileAccess.Read))
                {
                    _ = fs.Seek(parentEntry.Offset, SeekOrigin.Begin);
                    msndBuf = new byte[parentEntry.Size];
                    _ = fs.Read(msndBuf, 0, msndBuf.Length);
                }
                byte[] rebuilt = Msnd.ReplaceChunk(msndBuf, Path.GetExtension(chunkEntry.Path.Name).ToLowerInvariant(), newData);
                if (!string.IsNullOrEmpty(srcFolder))
                {
                    string outPath = Path.Combine(srcFolder, parentEntry.Path.Name);
                    _ = Directory.CreateDirectory(Path.GetDirectoryName(outPath) ?? srcFolder);
                    File.WriteAllBytes(outPath, rebuilt);
                }
                return rebuilt;
            }, ct).ConfigureAwait(false);
        }
        public async Task<byte[]> RebuildNestedFromFolderAsync(string folder, CancellationToken ct = default)
        {
            return await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                if (!Directory.Exists(folder))
                {
                    return null;
                }
                string mapperPath = Path.Combine(folder, "mapper.txt");
                if (File.Exists(mapperPath))
                {
                    Collection<Tuple<string, byte[]>> pairs = [];
                    foreach (string ln in File.ReadAllLines(mapperPath, Encoding.UTF8))
                    {
                        ct.ThrowIfCancellationRequested();
                        if (!ln.Contains('=', StringComparison.Ordinal))
                        {
                            continue;
                        }
                        string[] parts = ln.Split(separatorArray, 2);
                        string left = parts[0].Trim();
                        string right = parts[1].Trim();
                        string candidatePath = Path.Combine(folder, right);
                        if (Directory.Exists(candidatePath))
                        {
                            byte[] childBuf = RebuildNestedFromFolderAsync(candidatePath, ct).GetAwaiter().GetResult();
                            if (childBuf == null)
                            {
                                throw new InvalidOperationException($"Failed to rebuild nested archive at {candidatePath}");
                            }
                            pairs.Add(Tuple.Create(left, childBuf));
                            continue;
                        }
                        if (File.Exists(candidatePath))
                        {
                            pairs.Add(Tuple.Create(left, File.ReadAllBytes(candidatePath)));
                            continue;
                        }
                        string[] matches = Directory.GetFiles(folder, Path.GetFileName(right), SearchOption.AllDirectories);
                        if (matches?.Length > 0)
                        {
                            pairs.Add(Tuple.Create(left, File.ReadAllBytes(matches[0])));
                            continue;
                        }
                        throw new FileNotFoundException($"File not found: {candidatePath}");
                    }
                    return Dsarc.BuildFromPairs(pairs);
                }
                string baseName = Path.GetFileName(folder);
                string[] msndFiles = sourceArray.Select(ext => Directory.GetFiles(folder, $"*{ext}", SearchOption.TopDirectoryOnly).FirstOrDefault())
                    .ToArray();
                if (msndFiles.Any(f => f != null))
                {
                    Dictionary<string, byte[]> chunks = new(StringComparer.OrdinalIgnoreCase);
                    foreach (string ext in Msnd.MSNDORDER)
                    {
                        ct.ThrowIfCancellationRequested();
                        string exact = Path.Combine(folder, baseName + ext);
                        string? chosen = File.Exists(exact) ? exact : Directory.GetFiles(folder, $"*{ext}", SearchOption.TopDirectoryOnly).FirstOrDefault();
                        if (chosen == null)
                        {
                            throw new FileNotFoundException($"Missing expected MSND file {ext} in {folder}");
                        }
                        chunks[ext] = File.ReadAllBytes(chosen);
                    }
                    byte[]? txtBytes = null;
                    string txtPath = Path.Combine(folder, baseName + ".txt");
                    if (File.Exists(txtPath))
                    {
                        txtBytes = File.ReadAllBytes(txtPath);
                    }
                    return Msnd.Build(chunks, txtBytes);
                }
                string[] files = Directory.GetFiles(folder, "*", SearchOption.TopDirectoryOnly);
                return files.Length == 1 ? File.ReadAllBytes(files[0]) : throw new InvalidOperationException($"Cannot determine archive type to rebuild at {folder}");
            }, ct).ConfigureAwait(false);
        }
        public async Task NestedExtractBufferAsync(byte[] buf, string outdir, string baseLabel, IProgress<(int current, int total)>? progress = null, CancellationToken ct = default)
        {
            await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                if (buf == null || outdir == null)
                {
                    return;
                }
                _ = Directory.CreateDirectory(outdir);
                if (buf.Length >= 4 && buf.Take(4).SequenceEqual(Msnd.MAGICMSND))
                {
                    if (buf.Length >= Msnd.HDRMSND)
                    {
                        byte[] txt = new byte[4];
                        Array.Copy(buf, 44, txt, 0, Math.Min(4, buf.Length - 44));
                        File.WriteAllBytes(Path.Combine(outdir, baseLabel + ".txt"), txt);
                    }
                    Collection<Entry> children = Msnd.Parse(buf, baseLabel);
                    Dictionary<Tuple<string, string>, int> counters = new(_tupleComparer);
                    int total = children.Count;
                    for (int i = 0; i < children.Count; i++)
                    {
                        ct.ThrowIfCancellationRequested();
                        Entry child = children[i];
                        byte[] data = new byte[child.Size];
                        Array.Copy(buf, child.Offset, data, 0, child.Size);
                        bool childIsDsarc = data.Length >= 8 && data.Take(8).SequenceEqual(Dsarc.MAGICDSARC);
                        bool childIsMsnd = data.Length >= 4 && data.Take(4).SequenceEqual(Msnd.MAGICMSND);
                        if (childIsDsarc || childIsMsnd)
                        {
                            string folderName = UniqueOutName(Path.GetFileNameWithoutExtension(child.Path.Name), string.Empty, outdir, counters);
                            string childFolder = Path.Combine(outdir, folderName);
                            _ = Directory.CreateDirectory(childFolder);
                            NestedExtractBufferAsync(data, childFolder, Path.GetFileNameWithoutExtension(child.Path.Name), progress, ct).GetAwaiter().GetResult();
                            if (childIsMsnd)
                            {
                                string txtName = Path.GetFileNameWithoutExtension(child.Path.Name) + ".txt";
                                File.WriteAllBytes(Path.Combine(childFolder, txtName), Array.Empty<byte>());
                            }
                            else
                            {
                                File.WriteAllBytes(Path.Combine(childFolder, "mapper.txt"), Array.Empty<byte>());
                            }
                        }
                        else
                        {
                            string ext = Helpers.GuessExtByMagic(data, Path.GetExtension(child.Path.Name));
                            string finalName = UniqueOutName(Path.GetFileNameWithoutExtension(child.Path.Name), ext, outdir, counters);
                            File.WriteAllBytes(Path.Combine(outdir, finalName), data);
                        }
                        progress?.Report((i + 1, total));
                    }
                    return;
                }
                if (buf.Length >= 8 && buf.Take(8).SequenceEqual(Dsarc.MAGICDSARC))
                {
                    Collection<Entry> parsed = Dsarc.ParseFromBuffer(buf);
                    Dictionary<Tuple<string, string>, int> counters = new(_tupleComparer);
                    List<string> mappingLines = [];
                    int total = parsed.Count;
                    for (int i = 0; i < parsed.Count; i++)
                    {
                        ct.ThrowIfCancellationRequested();
                        Entry e = parsed[i];
                        byte[] data = new byte[e.Size];
                        Array.Copy(buf, e.Offset, data, 0, e.Size);
                        bool childIsDsarc = data.Length >= 8 && data.Take(8).SequenceEqual(Dsarc.MAGICDSARC);
                        bool childIsMsnd = data.Length >= 4 && data.Take(4).SequenceEqual(Msnd.MAGICMSND);
                        if (childIsDsarc || childIsMsnd)
                        {
                            string folderName = UniqueOutName(Path.GetFileNameWithoutExtension(e.Path.Name), string.Empty, outdir, counters);
                            string childFolder = Path.Combine(outdir, folderName);
                            _ = Directory.CreateDirectory(childFolder);
                            NestedExtractBufferAsync(data, childFolder, Path.GetFileNameWithoutExtension(e.Path.Name), progress, ct).GetAwaiter().GetResult();
                            mappingLines.Add($"{e.Path.Name}={folderName}");
                        }
                        else
                        {
                            string ext = Helpers.GuessExtByMagic(data, Path.GetExtension(e.Path.Name));
                            string finalName = UniqueOutName(Path.GetFileNameWithoutExtension(e.Path.Name), ext, outdir, counters);
                            File.WriteAllBytes(Path.Combine(outdir, finalName), data);
                            mappingLines.Add($"{e.Path.Name}={finalName}");
                        }
                        progress?.Report((i + 1, total));
                    }
                    File.WriteAllText(Path.Combine(outdir, "mapper.txt"), string.Join("\n", mappingLines), Encoding.UTF8);
                    return;
                }
                string fallbackName = UniqueOutName(baseLabel, Path.GetExtension(baseLabel), outdir, new Dictionary<Tuple<string, string>, int>(_tupleComparer));
                File.WriteAllBytes(Path.Combine(outdir, fallbackName), buf);
            }, ct).ConfigureAwait(false);
        }
        public async Task<Collection<Entry>> ParseDsarcFromBufferAsync(byte[] buf, CancellationToken ct = default)
        {
            return await Task.Run(() => Dsarc.ParseFromBuffer(buf), ct).ConfigureAwait(false);
        }
        public async Task<byte[]> BuildDsarcFromFolderAsync(string folder, CancellationToken ct = default)
        {
            return await Task.Run(() => BuildDsarcFromFolder(folder), ct).ConfigureAwait(false);
        }
        public async Task<byte[]> BuildMsndFromFolderAsync(string folder, CancellationToken ct = default)
        {
            return await Task.Run(() => BuildMsndFromFolder(folder), ct).ConfigureAwait(false);
        }
        public async Task WriteFileAsync(string path, byte[] data, CancellationToken ct = default)
        {
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentNullException(nameof(path));
            }
            await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                _ = Directory.CreateDirectory(Path.GetDirectoryName(path) ?? ".");
                File.WriteAllBytes(path, data);
            }, ct).ConfigureAwait(false);
        }
        public async Task<byte[]> ReadFileAsync(string path, CancellationToken ct = default)
        {
            return string.IsNullOrEmpty(path)
                ? throw new ArgumentNullException(nameof(path))
                : await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                return File.ReadAllBytes(path);
            }, ct).ConfigureAwait(false);
        }
        public async Task<byte[]> ReadRangeAsync(string path, long offset, int size, CancellationToken ct = default)
        {
            return string.IsNullOrEmpty(path)
                ? throw new ArgumentNullException(nameof(path))
                : await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                using FileStream fs = new(path, FileMode.Open, FileAccess.Read);
                _ = fs.Seek(offset, SeekOrigin.Begin);
                byte[] buf = new byte[size];
                int read = fs.Read(buf, 0, size);
                if (read != size)
                {
                    byte[] trimmed = new byte[read];
                    Array.Copy(buf, 0, trimmed, 0, read);
                    return trimmed;
                }
                return buf;
            }, ct).ConfigureAwait(false);
        }
        public async Task CopyFileToFolderAsync(string sourceFilePath, string destFolder, CancellationToken ct = default)
        {
            if (string.IsNullOrEmpty(sourceFilePath))
            {
                throw new ArgumentNullException(nameof(sourceFilePath));
            }
            if (string.IsNullOrEmpty(destFolder))
            {
                throw new ArgumentNullException(nameof(destFolder));
            }
            await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                _ = Directory.CreateDirectory(destFolder);
                string dest = Path.Combine(destFolder, Path.GetFileName(sourceFilePath));
                File.Copy(sourceFilePath, dest, true);
            }, ct).ConfigureAwait(false);
        }
        private byte[] BuildDsarcFromFolder(string folder)
        {
            if (!Directory.Exists(folder))
            {
                throw new DirectoryNotFoundException(folder);
            }
            string mapper = Path.Combine(folder, "mapper.txt");
            if (File.Exists(mapper))
            {
                Collection<Tuple<string, byte[]>> pairs = [];
                foreach (string ln in File.ReadAllLines(mapper, Encoding.UTF8))
                {
                    if (!ln.Contains('=', StringComparison.Ordinal))
                    {
                        continue;
                    }
                    string[] parts = ln.Split(separatorArray, 2);
                    string left = parts[0].Trim();
                    string right = parts[1].Trim();
                    string candidate = Path.Combine(folder, right);
                    if (Directory.Exists(candidate))
                    {
                        byte[] child = RebuildNestedFromFolderAsync(candidate).GetAwaiter().GetResult();
                        if (child == null)
                        {
                            throw new InvalidOperationException($"Failed to rebuild nested archive at {candidate}");
                        }
                        pairs.Add(Tuple.Create(left, child));
                        continue;
                    }
                    if (File.Exists(candidate))
                    {
                        pairs.Add(Tuple.Create(left, File.ReadAllBytes(candidate)));
                        continue;
                    }
                    string[] matches = Directory.GetFiles(folder, Path.GetFileName(right), SearchOption.AllDirectories);
                    if (matches?.Length > 0)
                    {
                        pairs.Add(Tuple.Create(left, File.ReadAllBytes(matches[0])));
                        continue;
                    }
                    throw new FileNotFoundException($"File not found: {candidate}");
                }
                return Dsarc.BuildFromPairs(pairs);
            }
            string[] files = Directory.GetFiles(folder, "*", SearchOption.AllDirectories);
            Collection<Tuple<string, byte[]>> list = [];
            foreach (string f in files)
            {
                string rel = GetRelativePath(folder, f);
                list.Add(Tuple.Create(rel, File.ReadAllBytes(f)));
            }
            return Dsarc.BuildFromPairs(list);
        }
        private static byte[] BuildMsndFromFolder(string folder)
        {
            if (!Directory.Exists(folder))
            {
                throw new DirectoryNotFoundException(folder);
            }
            string baseName = Path.GetFileName(folder);
            Dictionary<string, byte[]> chunks = new(StringComparer.OrdinalIgnoreCase);
            foreach (string ext in Msnd.MSNDORDER)
            {
                string exact = Path.Combine(folder, baseName + ext);
                string? chosen = File.Exists(exact) ? exact : Directory.GetFiles(folder, $"*{ext}", SearchOption.TopDirectoryOnly).FirstOrDefault();
                if (chosen == null)
                {
                    throw new FileNotFoundException($"Missing expected MSND file {ext} in {folder}");
                }
                chunks[ext] = File.ReadAllBytes(chosen);
            }
            byte[]? txtBytes = null;
            string txtPath = Path.Combine(folder, baseName + ".txt");
            if (File.Exists(txtPath))
            {
                txtBytes = File.ReadAllBytes(txtPath);
            }
            return Msnd.Build(chunks, txtBytes);
        }
        private static string UniqueOutName(string baseName, string ext, string outdir, Dictionary<Tuple<string, string>, int> counters)
        {
            Tuple<string, string> key = Tuple.Create(baseName, ext ?? string.Empty);
            if (!counters.TryGetValue(key, out int count))
            {
                count = 0;
            }
            count++;
            counters[key] = count;
            string candidate = count == 1
                ? (string.IsNullOrEmpty(ext) ? baseName : $"{baseName}{ext}")
                : (string.IsNullOrEmpty(ext) ? $"{baseName}_{count}" : $"{baseName}_{count}{ext}");
            string finalName = candidate;
            int extra = 1;
            while (File.Exists(Path.Combine(outdir, finalName)) || Directory.Exists(Path.Combine(outdir, finalName)))
            {
                finalName = string.IsNullOrEmpty(ext)
                    ? $"{baseName}_{count}_{extra++}"
                    : $"{baseName}_{count}_{extra++}{ext}";
            }
            return finalName;
        }
        private static string AppendDirectorySeparatorChar(string path)
        {
            return path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ? path : path + Path.DirectorySeparatorChar;
        }
        private static string GetRelativePath(string baseDir, string fullPath)
        {
            Uri baseUri = new(AppendDirectorySeparatorChar(baseDir));
            Uri fullUri = new(fullPath);
            return Uri.UnescapeDataString(baseUri.MakeRelativeUri(fullUri).ToString().Replace('/', Path.DirectorySeparatorChar));
        }
    }
}