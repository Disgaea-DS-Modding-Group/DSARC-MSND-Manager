using Avalonia;
using Avalonia.Collections;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

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

        public MainWindow() : this(new ArchiveService()) { }

        public MainWindow(IArchiveService archiveService)
        {
            _archiveService = archiveService ?? throw new ArgumentNullException(nameof(archiveService));
            InitializeComponent();

            if (!Design.IsDesignMode)
                WireUpEvents();
        }

        private void WireUpEvents()
        {
            var newMenu = this.FindControl<MenuItem>("NewMenu");
            var openMenu = this.FindControl<MenuItem>("OpenMenu");
            var saveMenu = this.FindControl<MenuItem>("SaveMenu");
            var saveAsMenu = this.FindControl<MenuItem>("SaveAsMenu");
            var exitMenu = this.FindControl<MenuItem>("ExitMenu");
            var tree = this.FindControl<TreeView>("TreeView");

            if (newMenu != null) newMenu.Click += (_, __) => NewArchive();
            if (openMenu != null) openMenu.Click += async (_, __) => await OpenArchiveAsync(_archiveService).ConfigureAwait(false);
            if (saveMenu != null) saveMenu.Click += async (_, __) => await SaveArchiveAsync().ConfigureAwait(false);
            if (saveAsMenu != null) saveAsMenu.Click += async (_, __) => await SaveAsAsync().ConfigureAwait(false);
            if (exitMenu != null) exitMenu.Click += (_, __) => Close();

            if (tree != null)
            {
                tree.SelectionChanged += TreeView_SelectionChanged;
                tree.AddHandler(InputElement.PointerReleasedEvent, TreeView_PointerReleased, RoutingStrategies.Tunnel);
            }

        }

        #region UI helpers

        private async Task<string?> SelectFolderAsync(string? title = null)
        {
            var dlg = new OpenFolderDialog();
            if (!string.IsNullOrEmpty(title)) dlg.Title = title;
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
            var dlg = new OpenFileDialog { AllowMultiple = false };
            try
            {
                var res = await dlg.ShowAsync(this);
                if (res != null && res.Length > 0) return res[0];
            }
            catch { }
            return null;
        }

        private async Task<string?> SelectSaveFileAsync(string filter = "All Files (*.*)|*.*")
        {
            var dlg = new SaveFileDialog();
            try
            {
                return await dlg.ShowAsync(this);
            }
            catch { }
            return null;
        }

        // In MainWindow.axaml.cs - Fix the dialog window creation
        private async Task ShowMessageDialogAsync(string text, string title = "")
        {
            await Dispatcher.UIThread.InvokeAsync(async () =>
            {
                var dialog = new Window
                {
                    Title = title,
                    Width = 480,
                    Height = 160,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    SizeToContent = SizeToContent.Manual
                };

                var panel = new StackPanel
                {
                    Margin = new Thickness(20),
                    VerticalAlignment = Avalonia.Layout.VerticalAlignment.Stretch
                };

                var textBlock = new TextBlock
                {
                    Text = text,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                    VerticalAlignment = Avalonia.Layout.VerticalAlignment.Stretch
                };

                var buttonPanel = new StackPanel
                {
                    Orientation = Avalonia.Layout.Orientation.Horizontal,
                    HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                    VerticalAlignment = Avalonia.Layout.VerticalAlignment.Bottom,
                    Margin = new Thickness(0, 20, 0, 0)
                };

                var okButton = new Button
                {
                    Content = "OK",
                    Width = 80,
                    IsDefault = true,
                    IsCancel = true
                };

                okButton.Click += (s, e) => dialog.Close();

                buttonPanel.Children.Add(okButton);
                panel.Children.Add(textBlock);
                panel.Children.Add(buttonPanel);
                dialog.Content = panel;

                await dialog.ShowDialog(this);
                AppendLog((string.IsNullOrEmpty(title) ? "" : title + ": ") + text);
            });
        }

        private Task ShowErrorAsync(string msg, string caption = "Error") => ShowMessageDialogAsync(msg, caption);
        private Task ShowWarningAsync(string msg, string caption = "Warning") => ShowMessageDialogAsync(msg, caption);

        private void AppendLog(string msg)
        {
            Dispatcher.UIThread.Post(() =>
            {
                var log = this.FindControl<TextBox>("LogTextBox");
                if (log != null)
                {
                    log.Text += $"{msg}\r\n";
                    log.CaretIndex = log.Text?.Length ?? 0;
                }
            });
        }

        private void UpdateProgress(int val, int total)
        {
            Dispatcher.UIThread.Post(() =>
            {
                var pb = this.FindControl<ProgressBar>("ProgressBar");
                if (pb == null) return;

                try
                {
                    int t = Math.Max(1, total);
                    pb.Maximum = t;
                    pb.Value = Math.Min(Math.Max(0, val), t);
                }
                catch (ArgumentOutOfRangeException)
                {
                    // Ignore progress bar range errors
                }
            });
        }

        private void SetStatus(string text)
        {
            Dispatcher.UIThread.Post(() =>
            {
                var lbl = this.FindControl<TextBlock>("StatusLabel");
                if (lbl != null) lbl.Text = text;
            });
        }

        #endregion

        #region Archive operations (port of your logic)

        private async void NewArchive()
        {
            try
            {
                string? file = await SelectSaveFileAsync().ConfigureAwait(false);
                if (string.IsNullOrEmpty(file)) return;
                archivePath = file;
                entries = new Collection<Entry>();
                srcFolder = null;
                filetype = archivePath.EndsWith(".msnd", StringComparison.OrdinalIgnoreCase) ? ArchiveType.MSND : ArchiveType.DSARC;
                archiveOpenedFromDisk = false;
                RefreshTree();
                SetStatus("New archive created");
                AppendLog($"Created new {filetype} archive: {Path.GetFileName(archivePath)}");
            }
            catch (Exception ex)
            {
                await ShowErrorAsync($"Failed to create new archive: {ex.Message}");
            }
        }

        private async Task OpenArchiveAsync(IArchiveService archiveService)
        {
            try
            {
                string? file = await SelectFileAsync().ConfigureAwait(false);
                if (file == null) return;
                archivePath = file;

                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                entries = await _archiveService.LoadArchiveAsync(archivePath, ct).ConfigureAwait(false);
                filetype = Detector.FromFile(archivePath);
                archiveOpenedFromDisk = true;
                RefreshTree();
                SetStatus($"Opened {Path.GetFileName(archivePath)}");
                AppendLog($"Opened {Path.GetFileName(archivePath)} as {filetype.ToString().ToUpper(CultureInfo.InvariantCulture)}");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Open archive cancelled.");
            }
            catch (IOException io)
            {
                await ShowErrorAsync($"Failed to open archive: {io.Message}");
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync($"Access denied: {ua.Message}");
            }
            catch (Exception ex)
            {
                await ShowErrorAsync($"Failed to open archive: {ex.Message}");
            }
            finally
            {
                _cts = null;
            }
        }

        // In MainWindow.axaml.cs - Fix the SaveArchiveAsync method to handle embedded MSND replacements
        private async Task SaveArchiveAsync()
        {
            try
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

                // Handle the case where srcFolder might be null but we have embedded MSND replacements
                if (srcFolder == null && !archiveOpenedFromDisk)
                {
                    // Check if we have any embedded MSND entries that might need source files
                    bool hasEmbeddedMsnd = entries.Any(e => e.IsMsnd && e.Children.Count > 0);
                    if (hasEmbeddedMsnd)
                    {
                        await ShowWarningAsync("Cannot save archive with embedded MSND replacements without a source folder. Please set a source folder first.");
                        return;
                    }

                    // For simple archives without embedded content, use a temporary approach
                    string? tempFolder = await SelectFolderAsync("Select temporary folder for archive contents");
                    if (tempFolder == null) return;
                    srcFolder = tempFolder;
                }

                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                Progress<(int current, int total)> progress = new(t => UpdateProgress(t.current, t.total));

                await _archiveService.SaveArchiveAsync(archivePath, filetype!.Value, entries.ToList(),
                    srcFolder ?? string.Empty, progress, ct).ConfigureAwait(false);

                SetStatus(filetype == ArchiveType.MSND ? "MSND saved" : "DSARC saved");
                AppendLog($"{filetype.ToString().ToUpper(CultureInfo.InvariantCulture)} saved.");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Save cancelled.");
            }
            catch (IOException io)
            {
                await ShowErrorAsync($"Failed to save archive: {io.Message}");
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync($"Access denied: {ua.Message}");
            }
            catch (Exception ex)
            {
                await ShowErrorAsync($"Failed to save archive: {ex.Message}");
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
                var tree = this.FindControl<TreeView>("TreeView");
                if (tree == null) return;

                string rootText = archivePath != null ? Path.GetFileName(archivePath) : "[New Archive]";
                rootItem = new TreeViewItem { Header = rootText, DataContext = null, IsExpanded = true };

                // Create child items for the root
                var rootChildren = new AvaloniaList<TreeViewItem>();

                if (filetype == ArchiveType.DSARC)
                {
                    foreach (Entry e in entries)
                    {
                        if (e.IsMsnd && e.Children?.Count > 0)
                        {
                            var msNode = new TreeViewItem { Header = e.Path.Name, DataContext = e, IsExpanded = true };
                            var childList = new AvaloniaList<TreeViewItem>();
                            foreach (Entry c in e.Children)
                            {
                                childList.Add(new TreeViewItem { Header = c.Path.Name, DataContext = c });
                            }
                            // set the child items via ItemsSourceProperty (not ItemsProperty)
                            msNode.SetValue(ItemsControl.ItemsSourceProperty, childList);
                            rootChildren.Add(msNode);
                        }
                        else
                        {
                            rootChildren.Add(new TreeViewItem { Header = e.Path.Name, DataContext = e });
                        }
                    }
                }
                else if (filetype == ArchiveType.MSND)
                {
                    foreach (Entry e in entries)
                    {
                        rootChildren.Add(new TreeViewItem { Header = e.Path.Name, DataContext = e });
                    }
                }

                // Set the root item's children
                rootItem.SetValue(ItemsControl.ItemsSourceProperty, rootChildren);

                // Create the main items list with just the root
                var allItems = new AvaloniaList<TreeViewItem> { rootItem };

                // set items via ItemsSourceProperty (Items is read-only)
                tree.SetValue(ItemsControl.ItemsSourceProperty, allItems);

                // Update context menu visibility for current selection
                UpdateContextMenuVisibility(tree.SelectedItem);
            });
        }


        #endregion

        #region Tree & Context menu

        private void TreeView_SelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            var tree = this.FindControl<TreeView>("TreeView");
            if (tree == null) return;

            var selectedItem = tree.SelectedItem;
            UpdateContextMenuVisibility(selectedItem);
        }

        private void TreeView_PointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            if (e.InitialPressMouseButton != MouseButton.Right) return;

            var tree = this.FindControl<TreeView>("TreeView");
            if (tree == null) return;

            var selectedItem = tree.SelectedItem;
            if (selectedItem is TreeViewItem treeViewItem)
            {
                ShowContextMenu(treeViewItem);
                e.Handled = true;
            }
        }

        private void UpdateContextMenuVisibility(object? selectedItem)
        {
            // This method is no longer needed with dynamic context menu creation
            // The context menu items are created dynamically in ShowContextMenu method
        }

        private void ShowContextMenu(TreeViewItem node)
        {
            var menu = new ContextMenu();
            var items = new AvaloniaList<MenuItem>();
            bool isRoot = node == rootItem;

            if (isRoot)
            {
                var miImport = new MenuItem { Header = "Import Folder" };
                miImport.Click += async (_, __) => await ImportFolderAsync().ConfigureAwait(false);
                items.Add(miImport);

                var miExtractAll = new MenuItem { Header = "Extract All" };
                miExtractAll.Click += async (_, __) => await ExtractAllAsync().ConfigureAwait(false);
                items.Add(miExtractAll);

                if (filetype == ArchiveType.DSARC)
                {
                    var miNested = new MenuItem { Header = "Extract All (Nested)" };
                    miNested.Click += async (_, __) => await ExtractAllNestedRootAsync().ConfigureAwait(false);
                    items.Add(miNested);
                }
            }
            else
            {
                var nodeEntry = node.DataContext as Entry;
                bool nodeHasChildren = node.Items != null && node.Items.Cast<object>().Any();

                if (nodeHasChildren && nodeEntry != null && nodeEntry.IsMsnd)
                {
                    var mi = new MenuItem { Header = "Import Folder" };
                    mi.Click += async (_, __) => await ImportFolderToNodeAsync().ConfigureAwait(false);
                    items.Add(mi);

                    var mi2 = new MenuItem { Header = "Extract All" };
                    mi2.Click += async (_, __) => await ExtractAllFromNodeAsync().ConfigureAwait(false);
                    items.Add(mi2);

                    var mi3 = new MenuItem { Header = "Extract All Nested" };
                    mi3.Click += async (_, __) => await ExtractAllNestedFromNodeAsync().ConfigureAwait(false);
                    items.Add(mi3);
                }

                if (filetype == ArchiveType.DSARC)
                {
                    if (node.Parent == rootItem)
                    {
                        var mi = new MenuItem { Header = "Extract" };
                        mi.Click += async (_, __) => await ExtractItemAsync().ConfigureAwait(false);
                        items.Add(mi);
                        var rep = new MenuItem { Header = "Replace" };
                        rep.Click += async (_, __) => await ReplaceItemAsync().ConfigureAwait(false);
                        items.Add(rep);
                    }
                    else if (node.Parent != null)
                    {
                        var mi = new MenuItem { Header = "Extract (chunk)" };
                        mi.Click += async (_, __) => await ExtractChunkItemAsync().ConfigureAwait(false);
                        items.Add(mi);
                        var rep = new MenuItem { Header = "Replace (chunk)" };
                        rep.Click += async (_, __) => await ReplaceChunkItemAsync().ConfigureAwait(false);
                        items.Add(rep);
                    }
                }
                else if (filetype == ArchiveType.MSND)
                {
                    var mi = new MenuItem { Header = "Extract" };
                    mi.Click += async (_, __) => await ExtractItemAsync().ConfigureAwait(false);
                    items.Add(mi);
                    var rep = new MenuItem { Header = "Replace" };
                    rep.Click += async (_, __) => await ReplaceItemAsync().ConfigureAwait(false);
                    items.Add(rep);
                }
            }

            // use ItemsSourceProperty instead of assigning Items
            menu.SetValue(ItemsControl.ItemsSourceProperty, items);
            menu.PlacementTarget = node;
            menu.Open(node);
        }

        #endregion

        #region Import/Export/Replace methods (delegating to _archiveService)

        private async Task ImportFolderAsync()
        {
            string? folder = await SelectFolderAsync("Select folder to import").ConfigureAwait(false);
            if (folder == null) return;
            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                var ct = _cts.Token;
                var result = await _archiveService.InspectFolderForImportAsync(folder, ct).ConfigureAwait(false);
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
                var ct = _cts.Token;
                Progress<(int current, int total)> progress = new(t => UpdateProgress(t.current, t.total));
                var result = await _archiveService.ExtractAllAsync(archivePath, filetype.Value, dlgFolder, progress, ct).ConfigureAwait(false);
                AppendLog($"Starting Extract All -> {result.outBase}");
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

        private async Task ExtractItemAsync()
        {
            var tree = this.FindControl<TreeView>("TreeView");
            if (tree?.SelectedItem is not TreeViewItem node) { await ShowWarningAsync("Invalid selection"); return; }

            if (rootItem == null) { await ShowWarningAsync("Invalid selection"); return; }

            int idx = -1;
            if (rootItem.Items != null)
            {
                var list = rootItem.Items.Cast<TreeViewItem>().ToList();
                idx = list.IndexOf(node);
            }
            if (idx < 0 || idx >= entries.Count) { await ShowWarningAsync("Invalid selection"); return; }

            Entry e = entries[idx];
            string? dest = await SelectFolderAsync("Select folder to extract to").ConfigureAwait(false);
            if (dest == null) return;

            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                await _archiveService.ExtractItemAsync(archivePath!, filetype!.Value, e, dest, _cts.Token).ConfigureAwait(false);
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

        private async Task ReplaceItemAsync()
        {
            var tree = this.FindControl<TreeView>("TreeView");
            if (tree?.SelectedItem is not TreeViewItem node) { await ShowWarningAsync("Invalid selection"); return; }

            if (rootItem == null) { await ShowWarningAsync("Invalid selection"); return; }

            int idx = -1;
            if (rootItem.Items != null)
            {
                var list = rootItem.Items.Cast<TreeViewItem>().ToList();
                idx = list.IndexOf(node);
            }
            if (idx < 0 || idx >= entries.Count) { await ShowWarningAsync("Invalid selection"); return; }

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
                await _archiveService.CopyFileToFolderAsync(replacement, srcFolder!, CancellationToken.None).ConfigureAwait(false);
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

        private async Task ExtractChunkItemAsync()
        {
            var tree = this.FindControl<TreeView>("TreeView");
            if (tree?.SelectedItem is not TreeViewItem node) { await ShowWarningAsync("Invalid selection"); return; }

            var parent = node.Parent as TreeViewItem;
            if (parent?.DataContext is not Entry parentEntry) { await ShowWarningAsync("Invalid parent"); return; }
            if (node.DataContext is not Entry chunkEntry) { await ShowWarningAsync("Invalid chunk"); return; }
            string? dest = await SelectFolderAsync("Select folder to extract chunk to").ConfigureAwait(false);
            if (dest == null) return;

            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();
                await _archive_service_ExtractChunkAsync(archivePath!, parentEntry, chunkEntry, dest, _cts.Token).ConfigureAwait(false);
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

        // small adapter to call IArchiveService extract chunk
        private Task _archive_service_ExtractChunkAsync(string archivePath, Entry parentEntry, Entry chunkEntry, string dest, CancellationToken ct)
            => _archiveService.ExtractChunkItemAsync(archivePath, parentEntry, chunkEntry, dest, ct);

        // In MainWindow.axaml.cs - Fix the ReplaceChunkItemAsync method
        private async Task<byte[]> ReplaceChunkItemAsync()
        {
            var tree = this.FindControl<TreeView>("TreeView");
            if (tree?.SelectedItem is not TreeViewItem node)
            {
                await ShowWarningAsync("Invalid selection");
                return Array.Empty<byte>();
            }

            var parent = node.Parent as TreeViewItem;
            if (parent?.DataContext is not Entry parentEntry)
            {
                await ShowWarningAsync("Invalid parent");
                return Array.Empty<byte>();
            }
            if (node.DataContext is not Entry chunkEntry)
            {
                await ShowWarningAsync("Invalid chunk");
                return Array.Empty<byte>();
            }

            string? replacement = await SelectFileAsync().ConfigureAwait(false);
            if (replacement == null) return Array.Empty<byte>();

            try
            {
                _cts?.Cancel();
                _cts = new CancellationTokenSource();

                // Ensure we have a valid srcFolder for the replacement
                if (string.IsNullOrEmpty(srcFolder))
                {
                    // If no srcFolder is set, use the directory of the replacement file
                    srcFolder = Path.GetDirectoryName(replacement);
                    if (string.IsNullOrEmpty(srcFolder))
                    {
                        await ShowErrorAsync("Cannot determine source folder for replacement");
                        return Array.Empty<byte>();
                    }
                }

                byte[] rebuilt = await _archiveService.ReplaceChunkItemAsync(
                    archivePath!, parentEntry, chunkEntry, replacement, srcFolder, _cts.Token).ConfigureAwait(false);

                // Update the parent entry with the new MSND structure
                parentEntry.Children.Clear();
                foreach (Entry child in Msnd.Parse(rebuilt, Path.GetFileNameWithoutExtension(parentEntry.Path.Name)))
                    parentEntry.Children.Add(child);

                RefreshTree();
                SetStatus("File replaced - use Save");
                AppendLog($"Rebuilt embedded MSND {parentEntry.Path.Name} after replacing {Path.GetExtension(chunkEntry.Path.Name)}");
                return rebuilt;
            }
            catch (OperationCanceledException)
            {
                AppendLog("Replace chunk cancelled.");
                return Array.Empty<byte>();
            }
            catch (IOException io)
            {
                await ShowErrorAsync(io.Message);
                return Array.Empty<byte>();
            }
            catch (UnauthorizedAccessException ua)
            {
                await ShowErrorAsync(ua.Message);
                return Array.Empty<byte>();
            }
            finally
            {
                _cts = null;
            }
        }

        private async Task ExtractAllFromNodeAsync()
        {
            var tree = this.FindControl<TreeView>("TreeView");
            if (tree?.SelectedItem is not TreeViewItem node) { await ShowWarningAsync("Invalid selection"); return; }

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

        private async Task ImportFolderToNodeAsync()
        {
            var tree = this.FindControl<TreeView>("TreeView");
            if (tree?.SelectedItem is not TreeViewItem node) { await ShowWarningAsync("Invalid selection"); return; }

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

        private async Task ExtractAllNestedFromNodeAsync()
        {
            var tree = this.FindControl<TreeView>("TreeView");
            if (tree?.SelectedItem is not TreeViewItem node) { await ShowWarningAsync("Invalid selection"); return; }

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
