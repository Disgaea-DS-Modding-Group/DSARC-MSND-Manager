using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using System.Collections;
using Avalonia;
using Avalonia.Input;
using Avalonia.VisualTree;

namespace Disgaea_DS_Manager
{
    public partial class MainWindow : Window
    {
        private string? archivePath;
        private Collection<Entry> entries = new();
        private string? srcFolder;
        private ArchiveType? filetype;
        private TreeViewItem? rootItem;
        private bool archiveOpenedFromDisk;
        private readonly IArchiveService _archiveService;
        private CancellationTokenSource? _cts;

        // Controls (found on InitializeComponent)
        private MenuItem? NewMenu;
        private MenuItem? OpenMenu;
        private MenuItem? SaveMenu;
        private MenuItem? SaveAsMenu;
        private MenuItem? ExitMenu;
        private TreeView? TreeView;
        private TextBox? LogTextBox;
        private TextBlock? StatusLabel;
        private ProgressBar? ProgressBar;

        public MainWindow() : this(new ArchiveService()) { }

        public MainWindow(IArchiveService archiveService)
        {
            _archiveService = archiveService ?? throw new ArgumentNullException(nameof(archiveService));
            InitializeComponent();

            // find named controls
            NewMenu = this.FindControl<MenuItem>("NewMenu");
            OpenMenu = this.FindControl<MenuItem>("OpenMenu");
            SaveMenu = this.FindControl<MenuItem>("SaveMenu");
            SaveAsMenu = this.FindControl<MenuItem>("SaveAsMenu");
            ExitMenu = this.FindControl<MenuItem>("ExitMenu");
            TreeView = this.FindControl<TreeView>("TreeView");
            LogTextBox = this.FindControl<TextBox>("LogTextBox");
            StatusLabel = this.FindControl<TextBlock>("StatusLabel");
            ProgressBar = this.FindControl<ProgressBar>("ProgressBar");

            if (!IsDesignTime())
            {
                WireUpEvents();
            }
        }

        private static bool IsDesignTime()
        {
            // basic detection — Avalonia doesn't use LicenseManager; keep simple check
            return Design.IsDesignMode;
        }

        private void WireUpEvents()
        {
            if (NewMenu != null) NewMenu.Click += (_, __) => NewArchive();
            if (OpenMenu != null) OpenMenu.Click += async (_, __) => await OpenArchiveAsync(_archiveService).ConfigureAwait(false);
            if (SaveMenu != null) SaveMenu.Click += async (_, __) => await SaveArchiveAsync().ConfigureAwait(false);
            if (SaveAsMenu != null) SaveAsMenu.Click += async (_, __) => await SaveAsAsync().ConfigureAwait(false);
            if (ExitMenu != null) ExitMenu.Click += (_, __) => this.Close();

            if (TreeView != null)
            {
                // selection changed: update selected node if needed
                TreeView.AddHandler(InputElement.PointerReleasedEvent, TreeView_PointerReleased, RoutingStrategies.Tunnel);
            }
        }

        #region Helpers: dialogs + UI thread helpers

        private async Task<string?> SelectFolderAsync(string? title = null)
        {
            var dlg = new OpenFolderDialog();
            if (!string.IsNullOrEmpty(title))
                dlg.Title = title;
            try
            {
                return await dlg.ShowAsync(this);
            }
            catch
            {
                return null;
            }
        }

        private async Task<string?> SelectFileAsync(string filter = "All Files (*.*)|*.*")
        {
            var dlg = new OpenFileDialog();
            dlg.AllowMultiple = false;
            // Avalonia's OpenFileDialog doesn't use WinForms-style filter string parsing automatically.
            // The consumer can pass a single-extension hint instead; we ignore for now.
            try
            {
                var result = await dlg.ShowAsync(this);
                if (result != null && result.Length > 0) return result[0];
                return null;
            }
            catch
            {
                return null;
            }
        }

        private async Task<string?> SelectSaveFileAsync(string filter = "All Files (*.*)|*.*")
        {
            var dlg = new SaveFileDialog();
            try
            {
                return await dlg.ShowAsync(this);
            }
            catch
            {
                return null;
            }
        }

        private Task ShowErrorAsync(string msg, string caption = "Error")
        {
            return ShowMessageDialogAsync(msg, caption);
        }

        private Task ShowWarningAsync(string msg, string caption = "Warning")
        {
            return ShowMessageDialogAsync(msg, caption);
        }

        private async Task ShowMessageDialogAsync(string text, string title = "")
        {
            // lightweight modal dialog implemented as a Window to avoid external packages.
            var win = new Window
            {
                Title = title,
                Width = 480,
                Height = 160,
                Content = new StackPanel
                {
                    Margin = new Thickness(12),
                    Children =
                    {
                        new TextBlock { Text = text, TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                        new StackPanel {
                            Orientation = Avalonia.Layout.Orientation.Horizontal,
                            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
                            Margin = new Thickness(0,12,0,0),
                            Children =
                            {
                                new Button { Content = "OK", IsDefault=true, IsCancel=true, Width = 90, HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right }
                            }
                        }
                    }
                }
            };

            // wire the button to close
            var btn = win.FindControl<Button>("") ?? win.GetVisualChildren().OfType<Button>().FirstOrDefault();
            if (btn != null)
            {
                btn.Click += (_, __) => win.Close();
            }

            try
            {
                await win.ShowDialog(this);
            }
            catch { /* ignore */ }
            AppendLog((string.IsNullOrEmpty(title) ? "" : title + ": ") + text);
        }

        private void AppendLog(string msg)
        {
            // Always run on UI thread
            Dispatcher.UIThread.Post(() =>
            {
                if (LogTextBox != null)
                {
                    // Append with newline
                    LogTextBox.Text += $"{msg}\r\n";
                    // scroll to end — TextBox doesn't have an API to scroll easily, but setting caret could help
                    LogTextBox.CaretIndex = LogTextBox.Text?.Length ?? 0;
                }
            });
        }

        private void UpdateProgress(int val, int total)
        {
            Dispatcher.UIThread.Post(() =>
            {
                if (ProgressBar == null) return;
                int t = Math.Max(1, total);
                ProgressBar.Maximum = t;
                // clamp
                var value = Math.Min(Math.Max(0, val), t);
                ProgressBar.Value = value;
            });
        }

        private void SetStatus(string text)
        {
            Dispatcher.UIThread.Post(() =>
            {
                if (StatusLabel != null) StatusLabel.Text = text;
            });
        }

        #endregion

        #region Archive operations (ported)
        private void NewArchive()
        {
            // show Save dialog to pick path
            _ = Dispatcher.UIThread.InvokeAsync(async () =>
            {
                string? file = await SelectSaveFileAsync().ConfigureAwait(false);
                if (string.IsNullOrEmpty(file))
                {
                    return;
                }
                archivePath = file;
                entries = new Collection<Entry>();
                srcFolder = null;
                filetype = archivePath.EndsWith(".msnd", StringComparison.OrdinalIgnoreCase) ? ArchiveType.MSND : ArchiveType.DSARC;
                archiveOpenedFromDisk = false;
                RefreshTree();
            });
        }

        private async Task OpenArchiveAsync(IArchiveService archiveService)
        {
            string? file = await SelectFileAsync().ConfigureAwait(false);
            if (file == null) return;
            archivePath = file;

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
                await ShowErrorAsync(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync(ua.Message);
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
                await ShowWarningAsync("No files in archive to save.");
                return;
            }
            if (srcFolder == null)
            {
                await ShowWarningAsync(archiveOpenedFromDisk ? "No files in archive to save." : "Import a folder via root context menu first.");
                return;
            }
            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                Progress<(int current, int total)> progress = new(t => UpdateProgress(t.current, t.total));
                await _archiveService.SaveArchiveAsync(archivePath, filetype.Value, entries, srcFolder, progress, ct).ConfigureAwait(false);
                SetStatus(filetype == ArchiveType.MSND ? "MSND saved" : "DSARC saved");
                AppendLog($"{filetype.ToString().ToUpper(CultureInfo.InvariantCulture)} saved.");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Save cancelled.");
            }
            catch (IOException io)
            {
                await ShowErrorAsync(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync(ua.Message);
            }
            finally
            {
                _cts = null;
                UpdateProgress(0, 1);
            }
        }

        private async Task SaveAsAsync()
        {
            string? file = await SelectSaveFileAsync().ConfigureAwait(false);
            if (file == null) return;
            archivePath = file;
            await SaveArchiveAsync().ConfigureAwait(false);
        }

        private void RefreshTree()
        {
            Dispatcher.UIThread.Post(() =>
            {
                if (TreeView == null) return;
                var items = new List<TreeViewItem>();

                string rootText = archivePath != null ? Path.GetFileName(archivePath) : "[New Archive]";
                rootItem = new TreeViewItem { Header = rootText, DataContext = null, IsExpanded = true };

                // Build children
                if (filetype == ArchiveType.DSARC)
                {
                    foreach (Entry e in entries)
                    {
                        if (e.IsMsnd && e.Children?.Count > 0)
                        {
                            var msNode = new TreeViewItem { Header = e.Path.Name, DataContext = e, IsExpanded = true };
                            foreach (Entry c in e.Children)
                            {
                                var childNode = new TreeViewItem { Header = c.Path.Name, DataContext = c };
                                msNode.Items = msNode.Items?.Concat(new[] { childNode }) ?? new[] { childNode };
                            }
                            rootItem.Items = rootItem.Items?.Concat(new[] { msNode }) ?? new[] { msNode };
                        }
                        else
                        {
                            var item = new TreeViewItem { Header = e.Path.Name, DataContext = e };
                            rootItem.Items = rootItem.Items?.Concat(new[] { item }) ?? new[] { item };
                        }
                    }
                }
                else if (filetype == ArchiveType.MSND)
                {
                    foreach (Entry e in entries)
                    {
                        var item = new TreeViewItem { Header = e.Path.Name, DataContext = e };
                        rootItem.Items = rootItem.Items?.Concat(new[] { item }) ?? new[] { item };
                    }
                }

                items.Add(rootItem);
                TreeView.Items = items;
            });
        }

        #endregion

        #region Tree interactions & context menu (simplified)
        private void TreeView_PointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            try
            {
                if (e.InitialPressMouseButton == MouseButton.Right)
                {
                    // We will use the currently selected item (TreeView.SelectedItem) when showing the context menu.
                    // Get the selected item and show a context menu suited to it.
                    var selected = TreeView?.SelectedItem;
                    // convert selected to TreeViewItem (if our Items are TreeViewItem objects) or otherwise match by header.
                    TreeViewItem? node = null;
                    if (selected is TreeViewItem tvi) node = tvi;
                    else
                    {
                        // our Items are TreeViewItem; SelectedItem may be the Header, so try to find by header
                        node = FindSelectedTreeViewItem();
                    }

                    if (node != null)
                    {
                        ShowContextMenu(node, e);
                        e.Handled = true;
                    }
                }
            }
            catch { }
        }

        private TreeViewItem? FindSelectedTreeViewItem()
        {
            // Try to find the TreeViewItem from tree by searching visual tree (a utility - may not always work depending on templates)
            if (TreeView == null) return null;
            var sel = TreeView.SelectedItem;
            if (sel == null) return null;

            foreach (var child in TreeView.GetVisualDescendants().OfType<TreeViewItem>())
            {
                if ((child.DataContext != null && child.DataContext.Equals(sel)) ||
                    (child.Header != null && child.Header.Equals(sel)))
                {
                    return child;
                }
            }
            return null;
        }

        private void ShowContextMenu(TreeViewItem node, RoutedEventArgs? pointerEvent)
        {
            // Build menu items based on whether node is root or an Entry
            var menu = new ContextMenu();
            var items = new List<MenuItem>();

            bool isRoot = node == rootItem;

            if (isRoot)
            {
                items.Add(new MenuItem { Header = "Import Folder", Command = Avalonia.Input.KeyGesture.Parse("")?.ToCommand() });
                items.Last().Click += async (_, __) => await ImportFolderAsync().ConfigureAwait(false);

                items.Add(new MenuItem { Header = "Extract All" });
                items.Last().Click += async (_, __) => await ExtractAllAsync().ConfigureAwait(false);

                if (filetype == ArchiveType.DSARC)
                {
                    items.Add(new MenuItem { Header = "Extract All (Nested)" });
                    items.Last().Click += async (_, __) => await ExtractAllNestedRootAsync().ConfigureAwait(false);
                }
            }
            else
            {
                var nodeEntry = node.DataContext as Entry;
                bool nodeHasChildren = node.Items != null && node.Items.Cast<object>().Any();

                if (nodeHasChildren && nodeEntry != null && nodeEntry.IsMsnd)
                {
                    var mi = new MenuItem { Header = "Import Folder" };
                    mi.Click += async (_, __) => await ImportFolderToNodeAsync(node).ConfigureAwait(false);
                    items.Add(mi);

                    var mi2 = new MenuItem { Header = "Extract All" };
                    mi2.Click += async (_, __) => await ExtractAllFromNodeAsync(node).ConfigureAwait(false);
                    items.Add(mi2);

                    var mi3 = new MenuItem { Header = "Extract All Nested" };
                    mi3.Click += async (_, __) => await ExtractAllNestedFromNodeAsync(node).ConfigureAwait(false);
                    items.Add(mi3);

                    items.Add(new MenuItem { Header = "-" });
                }

                if (filetype == ArchiveType.DSARC)
                {
                    if (node.Parent == rootItem)
                    {
                        var mi = new MenuItem { Header = "Extract" };
                        mi.Click += async (_, __) => await ExtractItemAsync(node).ConfigureAwait(false);
                        items.Add(mi);
                        var rep = new MenuItem { Header = "Replace" };
                        rep.Click += async (_, __) => await ReplaceItemAsync(node).ConfigureAwait(false);
                        items.Add(rep);
                    }
                    else if (node.Parent != null)
                    {
                        var mi = new MenuItem { Header = "Extract (chunk)" };
                        mi.Click += async (_, __) => await ExtractChunkItemAsync(node, (TreeViewItem)node.Parent).ConfigureAwait(false);
                        items.Add(mi);
                        var rep = new MenuItem { Header = "Replace (chunk)" };
                        rep.Click += async (_, __) => await ReplaceChunkItemAsync(node, (TreeViewItem)node.Parent).ConfigureAwait(false);
                        items.Add(rep);
                    }
                }
                else if (filetype == ArchiveType.MSND)
                {
                    var mi = new MenuItem { Header = "Extract" };
                    mi.Click += async (_, __) => await ExtractItemAsync(node).ConfigureAwait(false);
                    items.Add(mi);
                    var rep = new MenuItem { Header = "Replace" };
                    rep.Click += async (_, __) => await ReplaceItemAsync(node).ConfigureAwait(false);
                    items.Add(rep);
                }
            }

            menu.Items = items;
            // Open the context menu anchored to the tree view or node
            menu.PlacementTarget = node;
            menu.Open(node);
        }

        #endregion

        #region Import / Extract / Replace implementations (ported, minimal changes)

        private async Task ImportFolderAsync()
        {
            string? folder = await SelectFolderAsync("Select folder to import").ConfigureAwait(false);
            if (folder == null) return;
            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                ImportResult result = await _archiveService.InspectFolderForImportAsync(folder, ct).ConfigureAwait(false);
                if (result?.Entries == null || result.Entries.Count == 0)
                {
                    await ShowWarningAsync("No files found to import");
                    return;
                }
                entries = result.Entries;
                filetype = result.FileType;
                srcFolder = result.SourceFolder;
                RefreshTree();
                SetStatus("Folder imported");
                AppendLog("Folder imported (nested-aware).");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Import folder cancelled.");
            }
            catch (IOException io)
            {
                await ShowErrorAsync(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync(ua.Message);
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
                await ShowWarningAsync("Open archive first");
                return;
            }
            string? dlgFolder = await SelectFolderAsync("Choose output folder for Extract All").ConfigureAwait(false);
            if (dlgFolder == null) return;
            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                Progress<(int current, int total)> progress = new(t => UpdateProgress(t.current, t.total));
                (string outBase, List<string> mapper) = await _archive_service_ExtractAllAsync(archivePath, filetype.Value, dlgFolder, progress, ct).ConfigureAwait(false);
                AppendLog($"Starting Extract All -> {outBase}");
                SetStatus("Extract complete");
                AppendLog("Extract All complete.");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Extract All cancelled.");
            }
            catch (IOException io)
            {
                await ShowErrorAsync(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync(ua.Message);
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

        private async Task ExtractItemAsync(TreeViewItem node)
        {
            if (rootItem == null)
            {
                await ShowWarningAsync("Invalid selection");
                return;
            }

            // Determine index of node within rootItem children
            int idx = -1;
            if (rootItem.Items != null)
            {
                var list = rootItem.Items.Cast<TreeViewItem>().ToList();
                idx = list.IndexOf(node);
            }

            if (idx < 0 || idx >= entries.Count)
            {
                await ShowWarningAsync("Invalid selection");
                return;
            }
            Entry e = entries[idx];

            string? dest = await SelectFolderAsync("Select folder to extract to").ConfigureAwait(false);
            if (dest == null) return;

            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                await _archiveService.ExtractItemAsync(archivePath, filetype.Value, e, dest, _cts.Token).ConfigureAwait(false);
                SetStatus(string.Format(CultureInfo.InvariantCulture, "Extracted {0}", e.Path.Name));
                AppendLog($"Extracted {e.Path.Name} ({e.Size} bytes)");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Extract item cancelled.");
            }
            catch (IOException io)
            {
                await ShowErrorAsync(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync(ua.Message);
            }
            finally
            {
                _cts = null;
            }
        }

        private async Task ReplaceItemAsync(TreeViewItem node)
        {
            if (rootItem == null)
            {
                await ShowWarningAsync("Invalid selection");
                return;
            }
            int idx = -1;
            if (rootItem.Items != null)
            {
                var list = rootItem.Items.Cast<TreeViewItem>().ToList();
                idx = list.IndexOf(node);
            }
            if (idx < 0 || idx >= entries.Count)
            {
                await ShowWarningAsync("Invalid selection");
                return;
            }
            string? replacement = await SelectFileAsync().ConfigureAwait(false);
            if (replacement == null) return;

            try
            {
                if (filetype == ArchiveType.MSND && !Msnd.MSNDORDER.Contains(Path.GetExtension(replacement).ToLowerInvariant()))
                {
                    await ShowWarningAsync("Replacement must be sseq/sbnk/swar");
                    return;
                }
                srcFolder ??= Path.GetDirectoryName(replacement);
                await _archiveService.CopyFileToFolderAsync(replacement, srcFolder).ConfigureAwait(false);
                entries[idx].Path = new FileInfo(Path.GetFileName(replacement));
                RefreshTree();
                SetStatus("File replaced");
            }
            catch (IOException io)
            {
                await ShowErrorAsync(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync(ua.Message);
            }
        }

        private async Task ExtractChunkItemAsync(TreeViewItem node, TreeViewItem parent)
        {
            if (parent.DataContext is not Entry parentEntry)
            {
                await ShowWarningAsync("Invalid parent");
                return;
            }
            if (node.DataContext is not Entry chunkEntry)
            {
                await ShowWarningAsync("Invalid chunk");
                return;
            }
            string? dest = await SelectFolderAsync("Select folder to extract chunk to").ConfigureAwait(false);
            if (dest == null) return;
            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                await _archiveService.ExtractChunkItemAsync(archivePath, parentEntry, chunkEntry, dest, _cts.Token).ConfigureAwait(false);
                SetStatus(string.Format(CultureInfo.InvariantCulture, "Extracted {0}", chunkEntry.Path.Name));
                AppendLog($"Extracted {chunkEntry.Path.Name} ({chunkEntry.Size} bytes)");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Extract chunk cancelled.");
            }
            catch (IOException io)
            {
                await ShowErrorAsync(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync(ua.Message);
            }
            finally
            {
                _cts = null;
            }
        }

        private async Task ReplaceChunkItemAsync(TreeViewItem node, TreeViewItem parent)
        {
            if (parent.DataContext is not Entry parentEntry)
            {
                await ShowWarningAsync("Invalid parent");
                return;
            }
            if (node.DataContext is not Entry chunkEntry)
            {
                await ShowWarningAsync("Invalid chunk");
                return;
            }
            string? replacement = await SelectFileAsync().ConfigureAwait(false);
            if (replacement == null) return;

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
                SetStatus("File replaced - use Save");
                AppendLog($"Rebuilt embedded MSND {parentEntry.Path.Name} after replacing {Path.GetExtension(chunkEntry.Path.Name)}");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Replace chunk cancelled.");
            }
            catch (IOException io)
            {
                await ShowErrorAsync(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync(ua.Message);
            }
            finally
            {
                _cts = null;
            }
        }

        private async Task ExtractAllFromNodeAsync(TreeViewItem node)
        {
            if (node.DataContext is not Entry entry || !entry.IsMsnd)
            {
                await ShowWarningAsync("Selection is not an embedded archive");
                return;
            }
            string? dest = await SelectFolderAsync("Select root folder for extraction").ConfigureAwait(false);
            if (dest == null) return;

            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                string outDir = Path.Combine(dest, Path.GetFileNameWithoutExtension(entry.Path.Name));
                Directory.CreateDirectory(outDir);
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
                SetStatus("Extract complete");
                AppendLog("Extract complete.");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Nested extract cancelled.");
            }
            catch (IOException io)
            {
                await ShowErrorAsync(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync(ua.Message);
            }
            finally
            {
                _cts = null;
                UpdateProgress(0, 1);
            }
        }

        private async Task ImportFolderToNodeAsync(TreeViewItem node)
        {
            if (node.DataContext is not Entry entry || !entry.IsMsnd)
            {
                await ShowWarningAsync("Selection not an embedded archive");
                return;
            }
            string? folder = await SelectFolderAsync("Select folder containing msnd parts").ConfigureAwait(false);
            if (folder == null) return;

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
                SetStatus("Imported and staged");
                AppendLog($"Imported and rebuilt embedded archive {entry.Path.Name} (staged to {outPath}).");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Import to node cancelled.");
            }
            catch (IOException io)
            {
                await ShowErrorAsync(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync(ua.Message);
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
                await ShowWarningAsync("Open DSARC first");
                return;
            }
            string? dlg = await SelectFolderAsync("Select base folder for nested extract").ConfigureAwait(false);
            if (dlg == null) return;

            string expected = Path.GetFileNameWithoutExtension(archivePath);
            string outdir = ResolveSelectedFolderForExpected(dlg, expected, out bool found, out int matches);
            if (!found)
            {
                string candidate = Path.Combine(dlg, expected);
                if (File.Exists(candidate))
                {
                    await ShowErrorAsync(string.Format(CultureInfo.InvariantCulture, "File exists cannot create folder: {0}", expected));
                    return;
                }
                Directory.CreateDirectory(candidate);
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
                SetStatus("Nested extract complete");
                AppendLog("Nested extract complete.");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Nested root extract cancelled.");
            }
            catch (IOException io)
            {
                await ShowErrorAsync(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync(ua.Message);
            }
            finally
            {
                _cts = null;
                UpdateProgress(0, 1);
            }
        }

        private async Task ExtractAllNestedFromNodeAsync(TreeViewItem node)
        {
            if (node.DataContext is not Entry entry)
            {
                await ShowWarningAsync("Invalid selection");
                return;
            }
            string? dlg = await SelectFolderAsync("Select base folder for nested extract from node").ConfigureAwait(false);
            if (dlg == null) return;
            string expected = Path.GetFileNameWithoutExtension(entry.Path.Name);
            string outdir = ResolveSelectedFolderForExpected(dlg, expected, out bool found, out int matches);
            if (!found)
            {
                string candidate = Path.Combine(dlg, expected);
                if (File.Exists(candidate))
                {
                    await ShowErrorAsync(string.Format(CultureInfo.InvariantCulture, "File exists cannot create folder: {0}", expected));
                    return;
                }
                Directory.CreateDirectory(candidate);
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
                SetStatus("Nested extract complete");
                AppendLog("Nested extract complete.");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Nested extract from node cancelled.");
            }
            catch (IOException io)
            {
                await ShowErrorAsync(io.Message);
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync(ua.Message);
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

        private async Task _archive_service_NestedExtractBufferAsync(byte[] buf, string outdir, string baseLabel, IProgress<(int current, int total)> progress, CancellationToken ct)
        {
            await _archiveService.NestedExtractBufferAsync(buf, outdir, baseLabel, progress, ct).ConfigureAwait(false);
        }

        #endregion
    }
}
