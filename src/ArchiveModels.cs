using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
namespace Disgaea_DS_Manager
{
    public enum ArchiveType
    {
        DSARC,
        MSND
    }
    public class ImportResult
    {
        public ArchiveType FileType { get; set; }
        public Collection<Entry> Entries { get; } = [];
        public required string SourceFolder { get; set; }
    }
    public static class Helpers
    {
        public const int NAMESZ = 40;
        public static byte[] PadName(string n)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(n ?? string.Empty);
            if (bytes.Length > NAMESZ)
            {
                bytes = bytes.Take(NAMESZ).ToArray();
            }
            return bytes.Concat(new byte[NAMESZ - bytes.Length]).ToArray();
        }
        public static string GuessExtByMagic(byte[] data, string defaultExt)
        {
            ArgumentNullException.ThrowIfNull(data);
            if (data.Length >= 4)
            {
                if (data[0] == 0x53 && data[1] == 0x57 && data[2] == 0x41 && data[3] == 0x56)
                {
                    return ".swav";
                }
                if (data[0] == 0x53 && data[1] == 0x54 && data[2] == 0x52 && data[3] == 0x4D)
                {
                    return ".strm";
                }
            }
            return defaultExt;
        }
    }
    public class Entry
    {
        public FileInfo Path { get; set; }
        public int Size { get; set; }
        public int Offset { get; set; }
        public bool IsMsnd { get; set; }
        public Collection<Entry> Children { get; } = [];
        public Entry(FileInfo path, int size = 0, int offset = 0, bool isMsnd = false)
        {
            Path = path;
            Size = size;
            Offset = offset;
            IsMsnd = isMsnd;
        }
    }
    public class TupleComparer : IEqualityComparer<Tuple<string, string>>
    {
        public bool Equals(Tuple<string, string> x, Tuple<string, string> y)
        {
            return x is null
                ? throw new ArgumentNullException(nameof(x))
                : y is null
                ? throw new ArgumentNullException(nameof(y))
                : string.Equals(x.Item1, y.Item1, StringComparison.Ordinal)
                   && string.Equals(x.Item2, y.Item2, StringComparison.Ordinal);
        }
        public int GetHashCode(Tuple<string, string> obj)
        {
            ArgumentNullException.ThrowIfNull(obj);
            int h1 = StringComparer.Ordinal.GetHashCode(obj.Item1 ?? string.Empty);
            int h2 = StringComparer.Ordinal.GetHashCode(obj.Item2 ?? string.Empty);
            return h1 ^ h2;
        }
    }
    public static class Detector
    {
        private static readonly byte[] MAGICDSARC =
        {
            0x44, 0x53, 0x41, 0x52, 0x43, 0x20, 0x46, 0x4C
        };
        private static readonly byte[] MAGICMSND =
        {
            0x44, 0x53, 0x45, 0x51
        };
        public static ArchiveType FromFile(string path)
        {
            using FileStream fs = new(path, FileMode.Open, FileAccess.Read);
            byte[] magic8 = new byte[8];
            int read = fs.Read(magic8, 0, 8);
            return read < 4
                ? throw new InvalidDataException("Unable to read file header.")
                : magic8.Take(4).SequenceEqual(MAGICMSND)
                ? ArchiveType.MSND
                : read >= 8 && magic8.SequenceEqual(MAGICDSARC)
                ? ArchiveType.DSARC
                : throw new InvalidDataException("Unknown archive format (magic mismatch).");
        }
    }
    public static class Msnd
    {
        public const int HDRMSND = 48;
        public static readonly string[] MSNDORDER =
        {
            ".sseq",
            ".sbnk",
            ".swar"
        };
        public static readonly byte[] MAGICMSND =
        {
            0x44, 0x53, 0x45, 0x51
        };
        public static Collection<Entry> Parse(byte[] buf, string baseName)
        {
            ArgumentNullException.ThrowIfNull(buf);
            if (!buf.Take(4).SequenceEqual(MAGICMSND))
            {
                throw new InvalidDataException("Not an MSND archive.");
            }
            if (buf.Length < HDRMSND)
            {
                throw new InvalidDataException("MSND too small.");
            }
            int sseqOff = BitConverter.ToInt32(buf, 16) - 16;
            int sbnkOff = BitConverter.ToInt32(buf, 20);
            int swarOff = BitConverter.ToInt32(buf, 24);
            int sseqSz = BitConverter.ToInt32(buf, 32);
            int sbnkSz = BitConverter.ToInt32(buf, 36);
            int swarSz = BitConverter.ToInt32(buf, 40);
            Tuple<string, int, int>[] chunks = new[]
            {
                Tuple.Create("SSEQ", sseqOff, sseqSz),
                Tuple.Create("SBNK", sbnkOff, sbnkSz),
                Tuple.Create("SWAR", swarOff, swarSz)
            };
            foreach (Tuple<string, int, int> t in chunks)
            {
                if (t.Item2 < 0 || t.Item3 < 0 || t.Item2 + t.Item3 > buf.Length)
                {
                    throw new InvalidDataException($"{t.Item1} file exceeds bounds.");
                }
            }
            Collection<Entry> result =
            [
                new Entry(new FileInfo($"{baseName}.sseq"), sseqSz, sseqOff),
                new Entry(new FileInfo($"{baseName}.sbnk"), sbnkSz, sbnkOff),
                new Entry(new FileInfo($"{baseName}.swar"), swarSz, swarOff)
            ];
            return result;
        }
        public static byte[] Build(Dictionary<string, byte[]> chunks, byte[]? txtBytes = null)
        {
            ArgumentNullException.ThrowIfNull(chunks);
            foreach (string ext in MSNDORDER)
            {
                if (!chunks.ContainsKey(ext))
                {
                    throw new InvalidOperationException($"Missing {ext}");
                }
            }
            byte[] sseq = chunks[".sseq"], sbnk = chunks[".sbnk"], swar = chunks[".swar"];
            int sseqOff = HDRMSND;
            int sbnkOff = sseqOff + sseq.Length;
            int swarOff = sbnkOff + sbnk.Length;
            byte[] hdr = new byte[HDRMSND];
            Array.Copy(MAGICMSND, 0, hdr, 0, 4);
            Array.Copy(BitConverter.GetBytes(sseqOff + 16), 0, hdr, 16, 4);
            Array.Copy(BitConverter.GetBytes(sbnkOff), 0, hdr, 20, 4);
            Array.Copy(BitConverter.GetBytes(swarOff), 0, hdr, 24, 4);
            Array.Copy(BitConverter.GetBytes(sseq.Length), 0, hdr, 32, 4);
            Array.Copy(BitConverter.GetBytes(sbnk.Length), 0, hdr, 36, 4);
            Array.Copy(BitConverter.GetBytes(swar.Length), 0, hdr, 40, 4);
            if (txtBytes != null)
            {
                if (txtBytes.Length != 4)
                {
                    throw new ArgumentException("txt must be 4 bytes", nameof(txtBytes));
                }
                Array.Copy(txtBytes, 0, hdr, 44, 4);
            }
            using MemoryStream ms = new();
            using BinaryWriter bw = new(ms);
            bw.Write(hdr);
            bw.Write(sseq);
            bw.Write(sbnk);
            bw.Write(swar);
            return ms.ToArray();
        }
        public static byte[] ReplaceChunk(byte[] msndBuf, string ext, byte[] newData)
        {
            ArgumentNullException.ThrowIfNull(msndBuf);
            ArgumentNullException.ThrowIfNull(ext);
            ArgumentNullException.ThrowIfNull(newData);
            Collection<Entry> entries = Parse(msndBuf, "temp");
            Dictionary<string, Entry> map = new(StringComparer.OrdinalIgnoreCase)
            {
                { ".sseq", entries[0] },
                { ".sbnk", entries[1] },
                { ".swar", entries[2] }
            };
            byte[] sseq_b = msndBuf.Skip(map[".sseq"].Offset).Take(map[".sseq"].Size).ToArray();
            byte[] sbnk_b = msndBuf.Skip(map[".sbnk"].Offset).Take(map[".sbnk"].Size).ToArray();
            byte[] swar_b = msndBuf.Skip(map[".swar"].Offset).Take(map[".swar"].Size).ToArray();
            switch (ext.ToLowerInvariant())
            {
                case ".sseq":
                    sseq_b = newData;
                    break;
                case ".sbnk":
                    sbnk_b = newData;
                    break;
                case ".swar":
                    swar_b = newData;
                    break;
                default:
                    throw new ArgumentException("Unsupported file", nameof(ext));
            }
            byte[]? txt = (msndBuf.Length >= HDRMSND) ? msndBuf.Skip(44).Take(4).ToArray() : null;
            return Build(new Dictionary<string, byte[]>
            {
                { ".sseq", sseq_b },
                { ".sbnk", sbnk_b },
                { ".swar", swar_b }
            }, txt);
        }
    }
    public static class Dsarc
    {
        public const int HDRDSARC = 16;
        public const int ENTRYINFOSZ = 8;
        public static readonly byte[] MAGICDSARC =
        {
            0x44, 0x53, 0x41, 0x52, 0x43, 0x20, 0x46, 0x4C
        };
        public const int VERSION = 1;
        private static readonly char[] separator = new[] { '=' };

        public static Collection<Entry> Parse(string path)
        {
            using FileStream fs = new(path, FileMode.Open, FileAccess.Read);
            using BinaryReader br = new(fs, Encoding.UTF8, true);
            if (!br.ReadBytes(8).SequenceEqual(MAGICDSARC))
            {
                throw new InvalidDataException("Not a DSARC.");
            }
            int count = br.ReadInt32();
            int version = br.ReadInt32();
            if (version != VERSION)
            {
                throw new NotSupportedException($"Unsupported DSARC version {version}");
            }
            _ = fs.Seek(HDRDSARC, SeekOrigin.Begin);
            long archiveSize = new FileInfo(path).Length;
            List<Tuple<string, int, int>> entriesMeta = [];
            for (int i = 0; i < count; i++)
            {
                byte[] raw = br.ReadBytes(Helpers.NAMESZ);
                if (raw.Length < Helpers.NAMESZ)
                {
                    throw new InvalidDataException("Corrupted entry name");
                }
                string name = Encoding.UTF8.GetString(raw).Split('\0')[0].Trim();
                if (string.IsNullOrEmpty(name))
                {
                    name = $"file_{i}";
                }
                byte[] info = br.ReadBytes(ENTRYINFOSZ);
                if (info.Length < ENTRYINFOSZ)
                {
                    throw new InvalidDataException("Corrupted entry info");
                }
                int size = BitConverter.ToInt32(info, 0);
                int offset = BitConverter.ToInt32(info, 4);
                if (offset + (long)size > archiveSize)
                {
                    throw new InvalidDataException($"{name} exceeds bounds");
                }
                entriesMeta.Add(Tuple.Create(name, size, offset));
            }
            Collection<Entry> outList = [];
            foreach (Tuple<string, int, int> t in entriesMeta)
            {
                Entry e = new(new FileInfo(t.Item1), t.Item2, t.Item3);
                _ = fs.Seek(t.Item3, SeekOrigin.Begin);
                byte[] magicAttempt = br.ReadBytes(4);
                if (magicAttempt.SequenceEqual(Msnd.MAGICMSND))
                {
                    _ = fs.Seek(t.Item3, SeekOrigin.Begin);
                    byte[] msndBuf = br.ReadBytes(t.Item2);
                    e.IsMsnd = true;
                    foreach (Entry child in Msnd.Parse(msndBuf, Path.GetFileNameWithoutExtension(t.Item1)))
                    {
                        e.Children.Add(child);
                    }
                }
                outList.Add(e);
            }
            return outList;
        }
        public static Collection<Entry> ParseFromBuffer(byte[] buf)
        {
            ArgumentNullException.ThrowIfNull(buf);
            using MemoryStream ms = new(buf);
            using BinaryReader br = new(ms, Encoding.UTF8, true);
            if (!br.ReadBytes(8).SequenceEqual(MAGICDSARC))
            {
                throw new InvalidDataException("Not a DSARC in buffer.");
            }
            int count = br.ReadInt32();
            int version = br.ReadInt32();
            if (version != VERSION)
            {
                throw new NotSupportedException($"Unsupported DSARC version {version}");
            }
            _ = ms.Seek(HDRDSARC, SeekOrigin.Begin);
            List<Tuple<string, int, int>> entriesMeta = [];
            for (int i = 0; i < count; i++)
            {
                byte[] raw = br.ReadBytes(Helpers.NAMESZ);
                if (raw.Length < Helpers.NAMESZ)
                {
                    throw new InvalidDataException("Corrupted entry name");
                }
                string name = Encoding.UTF8.GetString(raw).Split('\0')[0].Trim();
                if (string.IsNullOrEmpty(name))
                {
                    name = $"file_{i}";
                }
                byte[] info = br.ReadBytes(ENTRYINFOSZ);
                if (info.Length < ENTRYINFOSZ)
                {
                    throw new InvalidDataException("Corrupted entry info");
                }
                int size = BitConverter.ToInt32(info, 0);
                int offset = BitConverter.ToInt32(info, 4);
                if (offset + (long)size > buf.Length)
                {
                    throw new InvalidDataException($"{name} exceeds bounds in buffer");
                }
                entriesMeta.Add(Tuple.Create(name, size, offset));
            }
            Collection<Entry> outList = [];
            foreach (Tuple<string, int, int> t in entriesMeta)
            {
                Entry e = new(new FileInfo(t.Item1), t.Item2, t.Item3);
                _ = ms.Seek(t.Item3, SeekOrigin.Begin);
                byte[] magicAttempt = br.ReadBytes(4);
                if (magicAttempt.SequenceEqual(Msnd.MAGICMSND))
                {
                    _ = ms.Seek(t.Item3, SeekOrigin.Begin);
                    byte[] msndBuf = br.ReadBytes(t.Item2);
                    e.IsMsnd = true;
                    foreach (Entry child in Msnd.Parse(msndBuf, Path.GetFileNameWithoutExtension(t.Item1)))
                    {
                        e.Children.Add(child);
                    }
                }
                outList.Add(e);
            }
            return outList;
        }
        public static byte[] Build(Collection<Entry> entries, string srcDir, string? mappingPath = null)
        {
            ArgumentNullException.ThrowIfNull(entries);
            ArgumentNullException.ThrowIfNull(srcDir);
            Collection<Tuple<string, byte[]>> pairs = [];
            if (!string.IsNullOrEmpty(mappingPath) && File.Exists(mappingPath))
            {
                foreach (string ln in File.ReadAllLines(mappingPath, Encoding.UTF8))
                {
                    if (!ln.Contains('=', StringComparison.Ordinal))
                    {
                        continue;
                    }
                    string[] parts = ln.Split(separator, 2);
                    string left = parts[0].Trim();
                    string right = parts[1].Trim();
                    string sourceFile = Path.Combine(srcDir, right);
                    if (!File.Exists(sourceFile))
                    {
                        throw new FileNotFoundException($"File missing: {right}");
                    }
                    pairs.Add(Tuple.Create(left, File.ReadAllBytes(sourceFile)));
                }
            }
            else
            {
                foreach (Entry e in entries)
                {
                    string sourceFile = Path.Combine(srcDir, e.Path.ToString());
                    if (!File.Exists(sourceFile))
                    {
                        throw new FileNotFoundException($"Missing source file: {sourceFile}");
                    }
                    pairs.Add(Tuple.Create(e.Path.ToString(), File.ReadAllBytes(sourceFile)));
                }
            }
            return BuildFromPairs(pairs);
        }
        public static byte[] BuildFromPairs(Collection<Tuple<string, byte[]>> pairs)
        {
            ArgumentNullException.ThrowIfNull(pairs);
            int count = pairs.Count;
            using MemoryStream ms = new();
            using BinaryWriter bw = new(ms, Encoding.UTF8, true);
            bw.Write(MAGICDSARC);
            bw.Write(count);
            bw.Write(VERSION);
            int offset = HDRDSARC + (count * (Helpers.NAMESZ + ENTRYINFOSZ));
            foreach (Tuple<string, byte[]> p in pairs)
            {
                byte[] nameBytes = Helpers.PadName(p.Item1);
                bw.Write(nameBytes);
                bw.Write(p.Item2.Length);
                bw.Write(offset);
                offset += p.Item2.Length;
            }
            foreach (Tuple<string, byte[]> p in pairs)
            {
                bw.Write(p.Item2);
            }
            return ms.ToArray();
        }
    }
}