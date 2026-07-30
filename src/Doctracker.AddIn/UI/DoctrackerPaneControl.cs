using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Doctracker.AddIn.Excel;
using Doctracker.AddIn.Infrastructure;
using Doctracker.Core.Geometry;
using Doctracker.Core.Models;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Doctracker.AddIn.UI
{
    internal sealed class DoctrackerPaneControl : UserControl
    {
        private readonly ExcelInterop.Application application;
        private readonly WorkbookProjectContext context;
        private readonly ExcelCellGateway cells;
        private readonly IOcrEngine ocr;
        private readonly DocumentCanvas canvas;
        private readonly ListBox documents;
        private readonly Label status;

        public DoctrackerPaneControl(ExcelInterop.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            context = new WorkbookProjectContext();
            cells = new ExcelCellGateway(application);
            ocr = new TesseractOcrEngine();

            Dock = DockStyle.Fill;
            Size = new Size(760, 700);
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9F);

            var header = BuildHeader();
            var snipToolbar = BuildSnipToolbar();

            documents = new ListBox
            {
                Dock = DockStyle.Fill,
                DisplayMember = "OriginalName",
                BorderStyle = BorderStyle.None,
                IntegralHeight = false
            };
            documents.SelectedIndexChanged += Documents_SelectedIndexChanged;

            var importButton = new Button
            {
                Text = "Ajouter des pièces",
                Dock = DockStyle.Top,
                Height = 36,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(18, 28, 45),
                ForeColor = Color.White
            };
            importButton.FlatAppearance.BorderSize = 0;
            importButton.Click += (sender, args) => ImportDocuments();

            var documentPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8) };
            documentPanel.Controls.Add(documents);
            documentPanel.Controls.Add(importButton);

            canvas = new DocumentCanvas();
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                FixedPanel = FixedPanel.Panel1,
                Size = new Size(760, 600),
                Panel1MinSize = 150
            };
            split.SplitterDistance = 185;
            split.Panel1.Controls.Add(documentPanel);
            split.Panel2.Controls.Add(canvas);

            status = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 28,
                Text = "Prêt",
                Padding = new Padding(10, 6, 0, 0),
                ForeColor = Color.FromArgb(80, 88, 98),
                BackColor = Color.FromArgb(248, 249, 251)
            };

            Controls.Add(split);
            Controls.Add(status);
            Controls.Add(snipToolbar);
            Controls.Add(header);
        }

        public void RefreshProject()
        {
            try
            {
                EnsureProject();
                BindDocuments();
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        public void ImportDocuments()
        {
            try
            {
                EnsureProject();
                using (var dialog = new OpenFileDialog
                {
                    Title = "Ajouter des pièces au dossier Doctracker",
                    Filter = "Documents|*.pdf;*.png;*.jpg;*.jpeg;*.tif;*.tiff;*.bmp",
                    Multiselect = true
                })
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    foreach (var path in dialog.FileNames)
                    {
                        context.Importer.Import(context.State, path, Environment.UserName);
                    }
                }
                BindDocuments();
                SetStatus("Pièces ajoutées au dossier local.");
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        public async void CreateSnip(SnipType type)
        {
            try
            {
                EnsureProject();
                var document = SelectedDocument;
                if (document == null) throw new InvalidOperationException("Select a source document.");
                if (!canvas.HasSelection) throw new InvalidOperationException("Draw a zone on the document.");

                var target = cells.GetSingleTarget();
                var worksheet = (ExcelInterop.Worksheet)target.Worksheet;
                var address = target.Address[false, false, ExcelInterop.XlReferenceStyle.xlA1];
                var rectangle = canvas.GetNormalizedSelection();
                string recognized;

                SetStatus("Reconnaissance OCR en cours…");
                using (var crop = canvas.CropSelection())
                {
                    recognized = await Task.Run(() => ocr.Recognize(crop));
                }

                var snip = context.Snips.Create(
                    context.State,
                    document.Id,
                    canvas.CurrentPageNumber,
                    new NormalizedRectangle(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height),
                    type,
                    recognized,
                    worksheet.Name,
                    address,
                    Environment.UserName);

                cells.WriteSnip(target, snip, document);
                canvas.ClearSelection();
                SetStatus(type + " créé et lié à " + worksheet.Name + "!" + address + ".");
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        public async void MatchSelection()
        {
            try
            {
                EnsureProject();
                var range = cells.GetSelection();
                if (range.Cells.CountLarge > 500)
                    throw new InvalidOperationException("For this version, select at most 500 cells per matching run.");

                SetStatus("Indexation OCR des pièces non indexées…");
                var indexer = new DocumentIndexer(context.Store, ocr);
                await Task.Run(() => indexer.IndexMissing(context.State,
                    (name, page, count) => SetStatusThreadSafe(
                        $"Indexation : {name} — page {page}/{count}")));

                var matched = 0;
                foreach (ExcelInterop.Range cell in range.Cells)
                {
                    var query = Convert.ToString(cell.Value2);
                    if (string.IsNullOrWhiteSpace(query)) continue;
                    var candidate = context.Matcher.Find(context.State, query, 1).FirstOrDefault();
                    if (candidate == null || candidate.Score < 0.55) continue;

                    var document = context.State.Documents.First(item => item.Id == candidate.DocumentId);
                    var worksheet = (ExcelInterop.Worksheet)cell.Worksheet;
                    var snip = context.Snips.Create(
                        context.State,
                        document.Id,
                        candidate.PageNumber,
                        new NormalizedRectangle(0, 0, 1, 1),
                        SnipType.Text,
                        query,
                        worksheet.Name,
                        cell.Address[false, false, ExcelInterop.XlReferenceStyle.xlA1],
                        Environment.UserName);
                    snip.Comment = "Rapprochement automatique, score " +
                                   candidate.Score.ToString("P0") + ". " + candidate.Evidence;
                    context.Store.Save(context.State);
                    cells.AttachProof(cell, snip, document);
                    matched++;
                }

                SetStatus($"{matched} cellule(s) rapprochée(s). Vérification humaine requise.");
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        public bool TryNavigateFromCell(ExcelInterop.Range target)
        {
            try
            {
                EnsureProject();
                var snipId = cells.GetSnipId(target);
                if (string.IsNullOrWhiteSpace(snipId)) return false;
                return NavigateToSnip(snipId);
            }
            catch (Exception exception)
            {
                ShowError(exception);
                return false;
            }
        }

        public void NavigateFromSelection()
        {
            try
            {
                var target = cells.GetSingleTarget();
                var snipId = cells.GetSnipId(target);
                if (string.IsNullOrWhiteSpace(snipId))
                    throw new InvalidOperationException("The selected cell has no Doctracker proof.");
                NavigateToSnip(snipId);
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        public void ReviewSelection()
        {
            try
            {
                EnsureProject();
                var target = cells.GetSingleTarget();
                var snipId = cells.GetSnipId(target);
                var snip = context.State.Snips.FirstOrDefault(item => item.Id == snipId);
                if (snip == null) throw new InvalidOperationException("The selected cell has no Doctracker proof.");

                using (var dialog = new ReviewDialog(snip))
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    context.Snips.SetReview(context.State, snip.Id, dialog.SelectedStatus,
                        dialog.ReviewComment, Environment.UserName);
                }

                var document = context.State.Documents.First(item => item.Id == snip.DocumentId);
                cells.AttachProof(target, snip, document);
                SetStatus("Statut de revue mis à jour : " + snip.Status + ".");
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private bool NavigateToSnip(string snipId)
        {
            var snip = context.State.Snips.FirstOrDefault(item => item.Id == snipId);
            if (snip == null) return false;
            var document = context.State.Documents.FirstOrDefault(item => item.Id == snip.DocumentId);
            if (document == null) return false;

            SelectDocument(document.Id);
            canvas.NavigateTo(context.Store.ResolveDocumentPath(document), snip);
            SetStatus(document.OriginalName + " — page " + snip.PageNumber + " — " + snip.Status);
            return true;
        }

        private Control BuildHeader()
        {
            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 58,
                BackColor = Color.FromArgb(255, 122, 0)
            };
            var title = new Label
            {
                Text = "DOCTRACKER",
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 15F, FontStyle.Bold),
                Location = new Point(12, 8),
                AutoSize = true
            };
            var subtitle = new Label
            {
                Text = "Preuve locale • OCR • Traçabilité",
                ForeColor = Color.FromArgb(255, 244, 230),
                Location = new Point(14, 36),
                AutoSize = true
            };
            header.Controls.Add(title);
            header.Controls.Add(subtitle);
            return header;
        }

        private Control BuildSnipToolbar()
        {
            var toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 43,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(6, 5, 6, 4),
                BackColor = Color.White,
                AutoScroll = true
            };
            AddToolbarButton(toolbar, "Texte", () => CreateSnip(SnipType.Text));
            AddToolbarButton(toolbar, "Nombre", () => CreateSnip(SnipType.Number));
            AddToolbarButton(toolbar, "Date", () => CreateSnip(SnipType.Date));
            AddToolbarButton(toolbar, "Somme", () => CreateSnip(SnipType.Sum));
            AddToolbarButton(toolbar, "Tableau", () => CreateSnip(SnipType.Table));
            AddToolbarButton(toolbar, "Matching", MatchSelection);
            AddToolbarButton(toolbar, "Ouvrir preuve", NavigateFromSelection);
            AddToolbarButton(toolbar, "Revoir", ReviewSelection);
            return toolbar;
        }

        private static void AddToolbarButton(Control parent, string text, Action action)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = true,
                Height = 29,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(248, 249, 251)
            };
            button.FlatAppearance.BorderColor = Color.FromArgb(220, 224, 230);
            button.Click += (sender, args) => action();
            parent.Controls.Add(button);
        }

        private void Documents_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                var document = SelectedDocument;
                if (document != null)
                {
                    canvas.LoadDocument(context.Store.ResolveDocumentPath(document));
                    SetStatus(document.OriginalName);
                }
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private DocumentRecord SelectedDocument => documents.SelectedItem as DocumentRecord;

        private void EnsureProject()
        {
            context.Ensure(application.ActiveWorkbook);
        }

        private void BindDocuments()
        {
            var selectedId = SelectedDocument?.Id;
            documents.DataSource = null;
            documents.DataSource = context.State.Documents.ToList();
            documents.DisplayMember = "OriginalName";
            if (selectedId != null) SelectDocument(selectedId);
            if (documents.SelectedIndex < 0 && documents.Items.Count > 0) documents.SelectedIndex = 0;
        }

        private void SelectDocument(string id)
        {
            for (var index = 0; index < documents.Items.Count; index++)
            {
                var document = documents.Items[index] as DocumentRecord;
                if (document != null && document.Id == id)
                {
                    documents.SelectedIndex = index;
                    return;
                }
            }
        }

        private void SetStatus(string message)
        {
            status.Text = message;
        }

        private void SetStatusThreadSafe(string message)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(SetStatus), message);
            }
            else
            {
                SetStatus(message);
            }
        }

        private void ShowError(Exception exception)
        {
            SetStatus("Erreur : " + exception.Message);
            MessageBox.Show(this, exception.Message, "Doctracker",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
