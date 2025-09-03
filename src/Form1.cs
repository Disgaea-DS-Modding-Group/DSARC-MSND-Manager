using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Versioning;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace Manager
{
    [SupportedOSPlatform("windows6.1")]
    internal partial class Form1 : Form
    {
        private string archivePath;
        private Collection<Entry> entries = [];
        private string? srcFolder;
        private ArchiveType? filetype;
        private TreeNode rootNode;
        private bool archiveOpenedFromDisk;
        private readonly IArchiveService _archiveService;
        private CancellationTokenSource? _cts;
        public Form1() : this(new ArchiveService()) { }
        public Form1(IArchiveService archiveService)
        {
            InitializeComponent();
            _archiveService = archiveService ?? throw new ArgumentNullException(nameof(archiveService));
            if (!IsDesignTime())
            {
                WireUpEvents();
            }
        }
        private void AppendLog(string msg)
        {
            if (InvokeRequired)
            {
                _ = BeginInvoke(() => AppendLog(msg));
                return;
            }
            logTextBox.AppendText($"{msg}\r\n");
        }
        private void ShowError(string msg, string caption = "Error")
        {
            if (InvokeRequired)
            {
                _ = BeginInvoke(() => ShowError(msg, caption));
                return;
            }
            _ = MessageBox.Show(msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            AppendLog($"ERROR: {msg}");
        }
        private void ShowWarning(string msg, string caption = "Warning")
        {
            if (InvokeRequired)
            {
                _ = BeginInvoke(() => ShowWarning(msg, caption));
                return;
            }
            _ = MessageBox.Show(msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            AppendLog($"WARNING: {msg}");
        }
        private static bool IsDesignTime()
        {
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime)
            {
                return true;
            }

            try
            {
                return System.Diagnostics.Process.GetCurrentProcess().ProcessName?.IndexOf("devenv", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch (InvalidOperationException)
            {
                return false;
            }
        }
        private void WireUpEvents()
        {
            newToolStripMenuItem.Click += (s, e) => NewArchive();
            openToolStripMenuItem.Click += async (s, e) => await OpenArchiveAsync(Get_archiveService()).ConfigureAwait(false);
            saveToolStripMenuItem.Click += async (s, e) => await SaveArchiveAsync().ConfigureAwait(false);
            saveAsToolStripMenuItem.Click += async (s, e) => await SaveAsAsync().ConfigureAwait(false);
            exitToolStripMenuItem.Click += (s, e) => Close();
            treeView1.NodeMouseClick += TreeView1_NodeMouseClick;
        }
        private static string? SelectFolder(string? description = null)
        {
            using FolderBrowserDialog dlg = new();
            if (!string.IsNullOrEmpty(description))
            {
                try
                {
                    dlg.Description = description;
                }
                catch (ArgumentException) { }
                catch (InvalidOperationException) { }
            }
            return dlg.ShowDialog() == DialogResult.OK ? dlg.SelectedPath : null;
        }
        private static string? SelectFile(string filter = "All Files (*.*)|*.*")
        {
            using OpenFileDialog dlg = new() { Filter = filter };
            return dlg.ShowDialog() == DialogResult.OK ? dlg.FileName : null;
        }
        private void NewArchive()
        {
            using (SaveFileDialog sfd = new() { Filter = Properties.Resources.Filter_DatMsnd })
            {
                if (sfd.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                archivePath = sfd.FileName;
            }
            entries = [];
            srcFolder = null;
            filetype = archivePath.EndsWith(".msnd", StringComparison.OrdinalIgnoreCase) ? ArchiveType.MSND : ArchiveType.DSARC;
            archiveOpenedFromDisk = false;
            RefreshTree();
        }
        private IArchiveService Get_archiveService()
        {
            return _archiveService;
        }
        private async Task OpenArchiveAsync(IArchiveService archiveService)
        {
            using (OpenFileDialog ofd = new() { Filter = Properties.Resources.Filter_DatMsnd })
            {
                if (ofd.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                archivePath = ofd.FileName;
            }
            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                entries = await _archiveService.LoadArchiveAsync(archivePath, ct).ConfigureAwait(false);
                filetype = Detector.FromFile(archivePath);
                archiveOpenedFromDisk = true;
                RefreshTree();
                AppendLog($"Opened {Path.GetFileName(archivePath)} as {filetype.ToString().ToUpper(CultureInfo.InvariantCulture)}");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Open archive cancelled.");
            }
            catch (IOException io)
            {
                ShowError(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                ShowError(ua.Message);
            }
            finally
            {
                _cts = null;
            }
        }
        private async Task SaveArchiveAsync()
        {
            if (archivePath == null)
            {
                await SaveAsAsync().ConfigureAwait(false);
                return;
            }
            if (entries?.Count == 0)
            {
                ShowWarning("No files in archive to save.");
                return;
            }
            if (srcFolder == null)
            {
                ShowWarning(archiveOpenedFromDisk ? "No files in archive to save." : "Import a folder via root context menu first.");
                return;
            }
            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                Progress<(int current, int total)> progress = new(t => UpdateProgress(t.current, t.total));
                await _archiveService.SaveArchiveAsync(archivePath, filetype.Value, entries, srcFolder, progress, ct).ConfigureAwait(false);
                statusLabel.Text = filetype == ArchiveType.MSND ? Properties.Resources.Status_MSND_Saved : Properties.Resources.Status_DSARC_Saved;
                AppendLog($"{filetype.ToString().ToUpper(CultureInfo.InvariantCulture)} saved.");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Save cancelled.");
            }
            catch (IOException io)
            {
                ShowError(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                ShowError(ua.Message);
            }
            finally
            {
                _cts = null;
                UpdateProgress(0, 1);
            }
        }
        private async Task SaveAsAsync()
        {
            using (SaveFileDialog sfd = new() { Filter = Properties.Resources.Filter_DatMsnd })
            {
                if (sfd.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                archivePath = sfd.FileName;
            }
            await SaveArchiveAsync().ConfigureAwait(false);
        }
        private void RefreshTree()
        {
            if (InvokeRequired)
            {
                _ = BeginInvoke(RefreshTree);
                return;
            }
            treeView1.Nodes.Clear();
            rootNode = new TreeNode(archivePath != null ? Path.GetFileName(archivePath) : "[New Archive]");
            _ = treeView1.Nodes.Add(rootNode);
            if (filetype == ArchiveType.DSARC)
            {
                foreach (Entry e in entries)
                {
                    if (e.IsMsnd && e.Children.Count > 0)
                    {
                        TreeNode msNode = new(e.Path.Name) { Tag = e };
                        foreach (Entry c in e.Children)
                        {
                            _ = msNode.Nodes.Add(new TreeNode(c.Path.Name) { Tag = c });
                        }
                        msNode.Expand();
                        _ = rootNode.Nodes.Add(msNode);
                    }
                    else
                    {
                        _ = rootNode.Nodes.Add(new TreeNode(e.Path.Name) { Tag = e });
                    }
                }
            }
            else if (filetype == ArchiveType.MSND)
            {
                foreach (Entry e in entries)
                {
                    _ = rootNode.Nodes.Add(new TreeNode(e.Path.Name) { Tag = e });
                }
            }
            rootNode.Expand();
        }
        private void TreeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            treeView1.SelectedNode = e.Node;
            if (e.Button == MouseButtons.Right)
            {
                ShowContextMenu(e.Node, e.Location);
            }
        }
        private void ShowContextMenu(TreeNode node, System.Drawing.Point location)
        {
            ContextMenuStrip menu = new();

            menu.Closed += (s, e) =>
            {
                try
                {
                    _ = BeginInvoke(() =>
                    {
                        try
                        {
                            menu.Dispose();
                        }
                        catch { }
                    });
                }
                catch { }
            };
            if (node == rootNode)
            {
                _ = menu.Items.Add(Properties.Resources.Menu_ImportFolder, null, async (s, e) => await ImportFolderAsync().ConfigureAwait(false));
                _ = menu.Items.Add(Properties.Resources.Menu_ExtractAll, null, async (s, e) => await ExtractAllAsync().ConfigureAwait(false));
                if (filetype == ArchiveType.DSARC)
                {
                    _ = menu.Items.Add(Properties.Resources.Menu_ExtractAllNested, null, async (s, e) => await ExtractAllNestedRootAsync().ConfigureAwait(false));
                }
            }
            else
            {
                bool nodeHasChildren = node.Nodes?.Count > 0;
                if (nodeHasChildren && node.Tag is Entry nodeEntry && nodeEntry.IsMsnd)
                {
                    _ = menu.Items.Add(Properties.Resources.Menu_ImportFolder, null, async (s, e) => await ImportFolderToNodeAsync(node).ConfigureAwait(false));
                    _ = menu.Items.Add(Properties.Resources.Menu_ExtractAll, null, async (s, e) => await ExtractAllFromNodeAsync(node).ConfigureAwait(false));
                    _ = menu.Items.Add(Properties.Resources.Menu_ExtractAllNested, null, async (s, e) => await ExtractAllNestedFromNodeAsync(node).ConfigureAwait(false));
                    _ = menu.Items.Add(new ToolStripSeparator());
                }
                if (filetype == ArchiveType.DSARC)
                {
                    if (node.Parent == rootNode)
                    {
                        _ = menu.Items.Add(Properties.Resources.Menu_Extract, null, async (s, e) => await ExtractItemAsync(node).ConfigureAwait(false));
                        _ = menu.Items.Add(Properties.Resources.Menu_Replace, null, async (s, e) => await ReplaceItemAsync(node).ConfigureAwait(false));
                    }
                    else if (node.Parent != null)
                    {
                        _ = menu.Items.Add(Properties.Resources.Menu_Extract, null, async (s, e) => await ExtractChunkItemAsync(node, node.Parent).ConfigureAwait(false));
                        _ = menu.Items.Add(Properties.Resources.Menu_Replace, null, async (s, e) => await ReplaceChunkItemAsync(node, node.Parent).ConfigureAwait(false));
                    }
                }
                else if (filetype == ArchiveType.MSND)
                {
                    _ = menu.Items.Add(Properties.Resources.Menu_Extract, null, async (s, e) => await ExtractItemAsync(node).ConfigureAwait(false));
                    _ = menu.Items.Add(Properties.Resources.Menu_Replace, null, async (s, e) => await ReplaceItemAsync(node).ConfigureAwait(false));
                }
            }
            menu.Show(treeView1, location);
        }
        private async Task ImportFolderAsync()
        {
            string folder = SelectFolder(Properties.Resources.Dialog_SelectFolderToImport);
            if (folder == null)
            {
                return;
            }

            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                ImportResult result = await _archiveService.InspectFolderForImportAsync(folder, ct).ConfigureAwait(false);
                if (result?.Entries == null || result.Entries.Count == 0)
                {
                    ShowWarning(Properties.Resources.Warning_NoFilesFoundToImport);
                    return;
                }
                entries = result.Entries;
                filetype = result.FileType;
                srcFolder = result.SourceFolder;
                RefreshTree();
                statusLabel.Text = Properties.Resources.Status_FolderImported;
                AppendLog("Folder imported (nested-aware).");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Import folder cancelled.");
            }
            catch (IOException io)
            {
                ShowError(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                ShowError(ua.Message);
            }
            finally
            {
                _cts = null;
            }
        }
        private async Task ExtractAllAsync()
        {
            if (archivePath == null || filetype == null)
            {
                ShowWarning(Properties.Resources.Warning_OpenArchiveFirst);
                return;
            }
            string dlgFolder = SelectFolder(Properties.Resources.Dialog_ChooseOutputFolderForExtractAll);
            if (dlgFolder == null)
            {
                return;
            }

            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                Progress<(int current, int total)> progress = new(t => UpdateProgress(t.current, t.total));
                (string outBase, List<string> mapper) = await _archive_service_ExtractAllAsync(archivePath, filetype.Value, dlgFolder, progress, ct).ConfigureAwait(false);
                AppendLog($"Starting Extract All -> {outBase}");
                statusLabel.Text = Properties.Resources.Status_ExtractComplete;
                AppendLog("Extract All complete.");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Extract All cancelled.");
            }
            catch (IOException io)
            {
                ShowError(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                ShowError(ua.Message);
            }
            finally
            {
                _cts = null;
                UpdateProgress(0, 1);
            }
        }
        private async Task<(string outBase, List<string> mapper)> _archive_service_ExtractAllAsync(string path, ArchiveType type, string dest, IProgress<(int current, int total)> progress, CancellationToken ct)
        {
            return await _archiveService.ExtractAllAsync(path, type, dest, progress, ct).ConfigureAwait(false);
        }
        private async Task ExtractItemAsync(TreeNode node)
        {
            int idx = rootNode.Nodes.IndexOf(node);
            if (idx < 0 || idx >= entries.Count)
            {
                ShowWarning(Properties.Resources.Warning_InvalidSelection);
                return;
            }
            Entry e = entries[idx];
            string dest = SelectFolder(Properties.Resources.Dialog_SelectFolderToExtractItemTo);
            if (dest == null)
            {
                return;
            }

            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                await _archiveService.ExtractItemAsync(archivePath, filetype.Value, e, dest, _cts.Token).ConfigureAwait(false);
                statusLabel.Text = string.Format(CultureInfo.InvariantCulture, Properties.Resources.Status_ExtractedFormat, e.Path.Name);
                AppendLog($"Extracted {e.Path.Name} ({e.Size} bytes)");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Extract item cancelled.");
            }
            catch (IOException io)
            {
                ShowError(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                ShowError(ua.Message);
            }
            finally
            {
                _cts = null;
            }
        }
        private async Task ReplaceItemAsync(TreeNode node)
        {
            int idx = rootNode.Nodes.IndexOf(node);
            if (idx < 0 || idx >= entries.Count)
            {
                ShowWarning(Properties.Resources.Warning_InvalidSelection);
                return;
            }
            string replacement = SelectFile("All Files (*.*)|*.*");
            if (replacement == null)
            {
                return;
            }

            try
            {
                if (filetype == ArchiveType.MSND && !Msnd.MSNDORDER.Contains(Path.GetExtension(replacement).ToLowerInvariant()))
                {
                    ShowWarning(Properties.Resources.Warning_ReplacementMustBeSseqSbnkSwar);
                    return;
                }
                srcFolder ??= Path.GetDirectoryName(replacement);
                await _archiveService.CopyFileToFolderAsync(replacement, srcFolder).ConfigureAwait(false);
                entries[idx].Path = new FileInfo(Path.GetFileName(replacement));
                RefreshTree();
                statusLabel.Text = Properties.Resources.Status_FileReplaced;
            }
            catch (IOException io)
            {
                ShowError(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                ShowError(ua.Message);
            }
        }
        private async Task ExtractChunkItemAsync(TreeNode node, TreeNode parent)
        {
            if (parent.Tag is not Entry parentEntry)
            {
                ShowWarning(Properties.Resources.Warning_InvalidParent);
                return;
            }
            if (node.Tag is not Entry chunkEntry)
            {
                ShowWarning(Properties.Resources.Warning_InvalidChunk);
                return;
            }
            string dest = SelectFolder(Properties.Resources.Dialog_SelectFolderToExtractChunkTo);
            if (dest == null)
            {
                return;
            }

            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                await _archiveService.ExtractChunkItemAsync(archivePath, parentEntry, chunkEntry, dest, _cts.Token).ConfigureAwait(false);
                statusLabel.Text = string.Format(CultureInfo.InvariantCulture, Properties.Resources.Status_ExtractedFormat, chunkEntry.Path.Name);
                AppendLog($"Extracted {chunkEntry.Path.Name} ({chunkEntry.Size} bytes)");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Extract chunk cancelled.");
            }
            catch (IOException io)
            {
                ShowError(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                ShowError(ua.Message);
            }
            finally
            {
                _cts = null;
            }
        }
        private async Task ReplaceChunkItemAsync(TreeNode node, TreeNode parent)
        {
            if (parent.Tag is not Entry parentEntry)
            {
                ShowWarning(Properties.Resources.Warning_InvalidParent);
                return;
            }
            if (node.Tag is not Entry chunkEntry)
            {
                ShowWarning(Properties.Resources.Warning_InvalidChunk);
                return;
            }
            string replacement = SelectFile("All Files (*.*)|*.*");
            if (replacement == null)
            {
                return;
            }

            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                byte[] rebuilt = await _archiveService.ReplaceChunkItemAsync(archivePath, parentEntry, chunkEntry, replacement, srcFolder, _cts.Token).ConfigureAwait(false);
                parentEntry.Children.Clear();
                foreach (Entry child in Msnd.Parse(rebuilt, Path.GetFileNameWithoutExtension(parentEntry.Path.Name)))
                {
                    parentEntry.Children.Add(child);
                }
                RefreshTree();
                statusLabel.Text = Properties.Resources.Status_FileReplacedUseSave;
                AppendLog($"Rebuilt embedded MSND {parentEntry.Path.Name} after replacing {Path.GetExtension(chunkEntry.Path.Name)}");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Replace chunk cancelled.");
            }
            catch (IOException io)
            {
                ShowError(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                ShowError(ua.Message);
            }
            finally
            {
                _cts = null;
            }
        }
        private async Task ExtractAllFromNodeAsync(TreeNode node)
        {
            if (node.Tag is not Entry entry || !entry.IsMsnd)
            {
                ShowWarning(Properties.Resources.Warning_SelectionNotEmbeddedArchive);
                return;
            }
            string dest = SelectFolder(Properties.Resources.Dialog_SelectRootFolderForExtraction);
            if (dest == null)
            {
                return;
            }

            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                string outDir = Path.Combine(dest, Path.GetFileNameWithoutExtension(entry.Path.Name));
                _ = Directory.CreateDirectory(outDir);
                byte[] msndBuf = await _archiveService.ReadRangeAsync(archivePath, entry.Offset, entry.Size, ct).ConfigureAwait(false);
                Progress<(int current, int total)> progress = new(t => UpdateProgress(t.current, t.total));
                await _archive_service_NestedExtractBufferAsync(msndBuf, outDir, Path.GetFileNameWithoutExtension(entry.Path.Name), progress, ct).ConfigureAwait(false);

                if (entry.Children.Count == 0)
                {
                    foreach (Entry child in Msnd.Parse(msndBuf, Path.GetFileNameWithoutExtension(entry.Path.Name)))
                    {
                        entry.Children.Add(child);
                    }
                }
                RefreshTree();
                statusLabel.Text = Properties.Resources.Status_ExtractComplete;
                AppendLog("Extract complete.");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Nested extract cancelled.");
            }
            catch (IOException io)
            {
                ShowError(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                ShowError(ua.Message);
            }
            finally
            {
                _cts = null;
                UpdateProgress(0, 1);
            }
        }
        private async Task ImportFolderToNodeAsync(TreeNode node)
        {
            if (node.Tag is not Entry entry || !entry.IsMsnd)
            {
                ShowWarning(Properties.Resources.Warning_SelectionNotEmbeddedArchive);
                return;
            }
            string folder = SelectFolder(Properties.Resources.Dialog_SelectFolderContainingMsndParts);
            if (folder == null)
            {
                return;
            }

            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                byte[] rebuilt = await _archiveService.BuildMsndFromFolderAsync(folder, ct).ConfigureAwait(false);
                srcFolder ??= folder;
                string outPath = Path.Combine(srcFolder, entry.Path.Name);
                await _archiveService.WriteFileAsync(outPath, rebuilt, ct).ConfigureAwait(false);
                entry.Children.Clear();
                foreach (Entry child in Msnd.Parse(rebuilt, Path.GetFileNameWithoutExtension(entry.Path.Name)))
                {
                    entry.Children.Add(child);
                }
                RefreshTree();
                statusLabel.Text = Properties.Resources.Status_ImportedAndStaged;
                AppendLog($"Imported and rebuilt embedded archive {entry.Path.Name} (staged to {outPath}).");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Import to node cancelled.");
            }
            catch (IOException io)
            {
                ShowError(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                ShowError(ua.Message);
            }
            finally
            {
                _cts = null;
            }
        }
        private async Task ExtractAllNestedRootAsync()
        {
            if (archivePath == null || filetype != ArchiveType.DSARC)
            {
                ShowWarning(Properties.Resources.Warning_OpenDsarcFirst);
                return;
            }
            string dlg = SelectFolder(Properties.Resources.Dialog_SelectBaseFolderForNestedExtract);
            if (dlg == null)
            {
                return;
            }

            string expected = Path.GetFileNameWithoutExtension(archivePath);
            string outdir = ResolveSelectedFolderForExpected(dlg, expected, out bool found, out int matches);
            if (!found)
            {
                string candidate = Path.Combine(dlg, expected);
                if (File.Exists(candidate))
                {
                    ShowError(string.Format(CultureInfo.InvariantCulture, Properties.Resources.Error_FileExistsCannotCreateFolder, expected));
                    return;
                }
                _ = Directory.CreateDirectory(candidate);
                AppendLog($"Created folder '{candidate}' for nested extract.");
                outdir = candidate;
                found = true;
                matches = 1;
            }
            if (matches > 1)
            {
                AppendLog($"Multiple candidate folders matched '{expected}'; using '{outdir}'.");
            }
            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                byte[] buf = await _archiveService.ReadFileAsync(archivePath, ct).ConfigureAwait(false);
                Progress<(int current, int total)> progress = new(t => UpdateProgress(t.current, t.total));
                await _archive_service_NestedExtractBufferAsync(buf, outdir, expected, progress, ct).ConfigureAwait(false);
                statusLabel.Text = Properties.Resources.Status_NestedExtractComplete;
                AppendLog("Nested extract complete.");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Nested root extract cancelled.");
            }
            catch (IOException io)
            {
                ShowError(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                ShowError(ua.Message);
            }
            finally
            {
                _cts = null;
                UpdateProgress(0, 1);
            }
        }
        private async Task ExtractAllNestedFromNodeAsync(TreeNode node)
        {
            if (node.Tag is not Entry entry)
            {
                ShowWarning(Properties.Resources.Warning_InvalidSelection);
                return;
            }
            string dlg = SelectFolder(Properties.Resources.Dialog_SelectBaseFolderForNestedExtractFromNode);
            if (dlg == null)
            {
                return;
            }

            string expected = Path.GetFileNameWithoutExtension(entry.Path.Name);
            string outdir = ResolveSelectedFolderForExpected(dlg, expected, out bool found, out int matches);
            if (!found)
            {
                string candidate = Path.Combine(dlg, expected);
                if (File.Exists(candidate))
                {
                    ShowError(string.Format(CultureInfo.InvariantCulture, Properties.Resources.Error_FileExistsCannotCreateFolder, expected));
                    return;
                }
                _ = Directory.CreateDirectory(candidate);
                AppendLog($"Created folder '{candidate}' for nested extract.");
                outdir = candidate;
                found = true;
                matches = 1;
            }
            if (matches > 1)
            {
                AppendLog($"Multiple candidate folders matched '{expected}'; using '{outdir}'.");
            }
            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                byte[] buf = await _archiveService.ReadRangeAsync(archivePath, entry.Offset, entry.Size, ct).ConfigureAwait(false);
                Progress<(int current, int total)> progress = new(t => UpdateProgress(t.current, t.total));
                await _archive_service_NestedExtractBufferAsync(buf, outdir, expected, progress, ct).ConfigureAwait(false);
                statusLabel.Text = Properties.Resources.Status_NestedExtractComplete;
                AppendLog("Nested extract complete.");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Nested extract from node cancelled.");
            }
            catch (IOException io)
            {
                ShowError(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                ShowError(ua.Message);
            }
            finally
            {
                _cts = null;
                UpdateProgress(0, 1);
            }
        }
        private string? ResolveSelectedFolderForExpected(string selectedPath, string expectedFolderName, out bool found, out int matches)
        {
            found = false;
            matches = 0;
            if (string.IsNullOrEmpty(selectedPath))
            {
                return null;
            }

            try
            {
                if (string.Equals(Path.GetFileName(selectedPath), expectedFolderName, StringComparison.OrdinalIgnoreCase))
                {
                    found = true;
                    matches = 1;
                    return selectedPath;
                }
                string direct = Path.Combine(selectedPath, expectedFolderName);
                if (Directory.Exists(direct))
                {
                    found = true;
                    matches = 1;
                    return direct;
                }
                string[] topMatches = Directory.GetDirectories(selectedPath, expectedFolderName, SearchOption.TopDirectoryOnly);
                if (topMatches?.Length > 0)
                {
                    found = true;
                    matches = topMatches.Length;
                    return topMatches[0];
                }
                string[] recMatches = Directory.GetDirectories(selectedPath, expectedFolderName, SearchOption.AllDirectories);
                if (recMatches?.Length > 0)
                {
                    found = true;
                    matches = recMatches.Length;
                    return recMatches[0];
                }
            }
            catch (IOException io)
            {
                AppendLog($"ResolveSelectedFolderForExpected error: {io.Message}");
            }
            catch (UnauthorizedAccessException ua)
            {
                AppendLog($"ResolveSelectedFolderForExpected error: {ua.Message}");
            }
            return null;
        }
        private void UpdateProgress(int val, int total)
        {
            if (!OperatingSystem.IsWindows())
            {
                return;
            }

            if (InvokeRequired)
            {
                _ = BeginInvoke(() => UpdateProgress(val, total));
                return;
            }
            try
            {
                int t = Math.Max(1, total);
                toolStripProgressBar1.Maximum = t;
                toolStripProgressBar1.Value = Math.Min(Math.Max(0, val), t);
            }
            catch (ArgumentOutOfRangeException) { }
        }
        private async Task _archive_service_NestedExtractBufferAsync(byte[] buf, string outdir, string baseLabel, IProgress<(int current, int total)> progress, CancellationToken ct)
        {
            await _archiveService.NestedExtractBufferAsync(buf, outdir, baseLabel, progress, ct).ConfigureAwait(false);
        }
    }
}