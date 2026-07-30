using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;
using Doctracker.AddIn.Infrastructure;
using Doctracker.Core.Models;
using PdfiumViewer;

namespace Doctracker.AddIn.UI
{
    internal sealed class DocumentCanvas : UserControl
    {
        private readonly PictureBox picture;
        private readonly Label pageLabel;
        private readonly Button previousButton;
        private readonly Button nextButton;
        private PdfDocument pdf;
        private Image currentImage;
        private string currentPath;
        private int pageIndex;
        private Point dragStart;
        private Point dragEnd;
        private bool dragging;
        private RectangleF? normalizedSelection;

        public DocumentCanvas()
        {
            Dock = DockStyle.Fill;
            BackColor = Color.FromArgb(239, 242, 246);

            var toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 36,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(4, 3, 4, 3),
                BackColor = Color.White
            };
            previousButton = new Button { Text = "◀", Width = 38, Height = 27 };
            nextButton = new Button { Text = "▶", Width = 38, Height = 27 };
            pageLabel = new Label
            {
                Text = "Aucun document",
                AutoSize = true,
                Padding = new Padding(8, 6, 0, 0)
            };
            previousButton.Click += (sender, args) => ShowPage(pageIndex - 1);
            nextButton.Click += (sender, args) => ShowPage(pageIndex + 1);
            toolbar.Controls.Add(previousButton);
            toolbar.Controls.Add(nextButton);
            toolbar.Controls.Add(pageLabel);

            picture = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.FromArgb(55, 62, 73),
                Cursor = Cursors.Cross
            };
            picture.MouseDown += Picture_MouseDown;
            picture.MouseMove += Picture_MouseMove;
            picture.MouseUp += Picture_MouseUp;
            picture.Paint += Picture_Paint;

            Controls.Add(picture);
            Controls.Add(toolbar);
        }

        public int CurrentPageNumber => pageIndex + 1;
        public bool HasDocument => currentImage != null;
        public bool HasSelection => normalizedSelection.HasValue;

        public void LoadDocument(string path)
        {
            DisposeDocument();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                throw new FileNotFoundException("The document cannot be found.", path);
            }

            try
            {
                currentPath = path;
                pageIndex = 0;
                normalizedSelection = null;

                if (string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    NativePdfiumLoader.EnsureLoaded();
                    pdf = PdfDocument.Load(path);
                }
                RenderCurrentPage();
            }
            catch
            {
                DisposeDocument();
                throw;
            }
        }

        public void NavigateTo(string path, SnipRecord snip)
        {
            if (!string.Equals(currentPath, path, StringComparison.OrdinalIgnoreCase))
            {
                LoadDocument(path);
            }
            ShowPage(Math.Max(0, snip.PageNumber - 1));
            normalizedSelection = new RectangleF(
                (float)snip.X, (float)snip.Y, (float)snip.Width, (float)snip.Height);
            picture.Invalidate();
        }

        public RectangleF GetNormalizedSelection()
        {
            if (!normalizedSelection.HasValue)
                throw new InvalidOperationException("Draw a zone on the document first.");
            return normalizedSelection.Value;
        }

        public Bitmap CropSelection()
        {
            if (currentImage == null) throw new InvalidOperationException("No document is open.");
            var selection = GetNormalizedSelection();
            var crop = Rectangle.FromLTRB(
                Math.Max(0, (int)Math.Floor(selection.Left * currentImage.Width)),
                Math.Max(0, (int)Math.Floor(selection.Top * currentImage.Height)),
                Math.Min(currentImage.Width, (int)Math.Ceiling(selection.Right * currentImage.Width)),
                Math.Min(currentImage.Height, (int)Math.Ceiling(selection.Bottom * currentImage.Height)));
            if (crop.Width < 2 || crop.Height < 2)
                throw new InvalidOperationException("The selected zone is too small.");

            var output = new Bitmap(crop.Width, crop.Height);
            var dpiX = currentImage.HorizontalResolution > 0 ? currentImage.HorizontalResolution : 96f;
            var dpiY = currentImage.VerticalResolution > 0 ? currentImage.VerticalResolution : 96f;
            output.SetResolution(dpiX, dpiY);
            using (var graphics = Graphics.FromImage(output))
            {
                graphics.DrawImage(currentImage,
                    new Rectangle(0, 0, crop.Width, crop.Height),
                    crop,
                    GraphicsUnit.Pixel);
            }
            return output;
        }

        public void ClearSelection()
        {
            normalizedSelection = null;
            picture.Invalidate();
        }

        private void ShowPage(int requestedIndex)
        {
            var pageCount = pdf == null ? (currentPath == null ? 0 : 1) : pdf.PageCount;
            if (requestedIndex < 0 || requestedIndex >= pageCount) return;
            pageIndex = requestedIndex;
            normalizedSelection = null;
            RenderCurrentPage();
        }

        private void RenderCurrentPage()
        {
            if (currentImage != null)
            {
                picture.Image = null;
                currentImage.Dispose();
                currentImage = null;
            }

            if (pdf != null)
            {
                currentImage = pdf.Render(
                    pageIndex, 1800, 2400, 144, 144,
                    PdfRenderFlags.Annotations | PdfRenderFlags.LcdText);
                pageLabel.Text = $"Page {pageIndex + 1} / {pdf.PageCount}";
            }
            else if (!string.IsNullOrWhiteSpace(currentPath))
            {
                using (var source = Image.FromFile(currentPath))
                {
                    currentImage = new Bitmap(source);
                }
                pageLabel.Text = "Page 1 / 1";
            }
            else
            {
                pageLabel.Text = "Aucun document";
            }

            picture.Image = currentImage;
            previousButton.Enabled = pageIndex > 0;
            nextButton.Enabled = pdf != null && pageIndex < pdf.PageCount - 1;
            picture.Invalidate();
        }

        private void Picture_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left || currentImage == null) return;
            var imageBounds = GetDisplayedImageBounds();
            if (!imageBounds.Contains(e.Location)) return;
            dragging = true;
            dragStart = Clamp(e.Location, imageBounds);
            dragEnd = dragStart;
            normalizedSelection = null;
        }

        private void Picture_MouseMove(object sender, MouseEventArgs e)
        {
            if (!dragging) return;
            dragEnd = Clamp(e.Location, GetDisplayedImageBounds());
            picture.Invalidate();
        }

        private void Picture_MouseUp(object sender, MouseEventArgs e)
        {
            if (!dragging) return;
            dragging = false;
            dragEnd = Clamp(e.Location, GetDisplayedImageBounds());
            var display = GetDisplayedImageBounds();
            if (display.Width <= 0 || display.Height <= 0)
            {
                normalizedSelection = null;
                picture.Invalidate();
                return;
            }
            var rectangle = NormalizeScreenRectangle(dragStart, dragEnd);
            if (rectangle.Width >= 4 && rectangle.Height >= 4)
            {
                normalizedSelection = new RectangleF(
                    (rectangle.Left - display.Left) / (float)display.Width,
                    (rectangle.Top - display.Top) / (float)display.Height,
                    rectangle.Width / (float)display.Width,
                    rectangle.Height / (float)display.Height);
            }
            picture.Invalidate();
        }

        private void Picture_Paint(object sender, PaintEventArgs e)
        {
            Rectangle rectangle;
            if (dragging)
            {
                rectangle = NormalizeScreenRectangle(dragStart, dragEnd);
            }
            else if (normalizedSelection.HasValue)
            {
                var display = GetDisplayedImageBounds();
                var selection = normalizedSelection.Value;
                rectangle = new Rectangle(
                    display.Left + (int)(selection.X * display.Width),
                    display.Top + (int)(selection.Y * display.Height),
                    (int)(selection.Width * display.Width),
                    (int)(selection.Height * display.Height));
            }
            else
            {
                return;
            }

            using (var fill = new SolidBrush(Color.FromArgb(45, 255, 122, 0)))
            using (var pen = new Pen(Color.FromArgb(255, 122, 0), 2f))
            {
                pen.DashStyle = DashStyle.Dash;
                e.Graphics.FillRectangle(fill, rectangle);
                e.Graphics.DrawRectangle(pen, rectangle);
            }
        }

        private Rectangle GetDisplayedImageBounds()
        {
            if (currentImage == null || picture.ClientSize.Width == 0 || picture.ClientSize.Height == 0)
                return Rectangle.Empty;

            var imageRatio = currentImage.Width / (double)currentImage.Height;
            var clientRatio = picture.ClientSize.Width / (double)picture.ClientSize.Height;
            if (imageRatio > clientRatio)
            {
                var height = (int)(picture.ClientSize.Width / imageRatio);
                if (height < 1) return Rectangle.Empty;
                return new Rectangle(0, (picture.ClientSize.Height - height) / 2,
                    picture.ClientSize.Width, height);
            }

            var width = (int)(picture.ClientSize.Height * imageRatio);
            if (width < 1) return Rectangle.Empty;
            return new Rectangle((picture.ClientSize.Width - width) / 2, 0,
                width, picture.ClientSize.Height);
        }

        private static Point Clamp(Point point, Rectangle bounds)
        {
            return new Point(
                Math.Max(bounds.Left, Math.Min(bounds.Right, point.X)),
                Math.Max(bounds.Top, Math.Min(bounds.Bottom, point.Y)));
        }

        private static Rectangle NormalizeScreenRectangle(Point first, Point second)
        {
            return Rectangle.FromLTRB(
                Math.Min(first.X, second.X),
                Math.Min(first.Y, second.Y),
                Math.Max(first.X, second.X),
                Math.Max(first.Y, second.Y));
        }

        private void DisposeDocument()
        {
            picture.Image = null;
            if (currentImage != null) currentImage.Dispose();
            if (pdf != null) pdf.Dispose();
            currentImage = null;
            pdf = null;
            currentPath = null;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) DisposeDocument();
            base.Dispose(disposing);
        }
    }
}
