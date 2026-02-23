using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace JeiAudit
{
    internal sealed class ReportViewerForm : Form
    {
        private sealed class GroupCounters
        {
            public int Total { get; set; }
            public int Pass { get; set; }
            public int Fail { get; set; }
            public int Skipped { get; set; }
            public int CountOnly { get; set; }
            public int Unsupported { get; set; }
            public int Executed => Pass + Fail;
            public int NotExecuted => Total - Executed;
            public int PassRateRounded => Executed > 0 ? (int)Math.Round((double)Pass * 100.0 / Executed, MidpointRounding.AwayFromZero) : 0;
        }

        private sealed class SectionSummaryRow
        {
            public string Heading { get; set; } = string.Empty;
            public string Section { get; set; } = string.Empty;
            public string Description { get; set; } = string.Empty;
            public int TotalChecks { get; set; }
            public int PassChecks { get; set; }
            public int FailChecks { get; set; }
            public int NotExecutedChecks { get; set; }
            public int PassRate { get; set; }
        }

        private sealed class CriticalFindingRow
        {
            public string Heading { get; set; } = string.Empty;
            public string Section { get; set; } = string.Empty;
            public string CheckName { get; set; } = string.Empty;
            public int FailedElements { get; set; }
            public int CandidateElements { get; set; }
            public string Reason { get; set; } = string.Empty;
        }

        private sealed class ParameterFillRateRow
        {
            public string ParameterName { get; set; } = string.Empty;
            public int CandidateElements { get; set; }
            public int EmptyElements { get; set; }
            public int FilledElements => Math.Max(0, CandidateElements - EmptyElements);
            public double FillPercent => CandidateElements > 0
                ? FilledElements * 100.0 / CandidateElements
                : 0.0;
        }

        private readonly AuditReportData _data;
        private readonly Dictionary<string, List<ModelCheckerFailedElementItem>> _failedByCheck;

        private readonly Label _titleValueLabel;
        private readonly Label _dateValueLabel;
        private readonly Label _authorValueLabel;
        private readonly Label _descriptionValueLabel;
        private readonly Label _modelTagLabel;
        private readonly Label _scoreLabel;
        private readonly Label _summaryLabel;
        private readonly Label _reportDateLabel;
        private readonly Label _modelPathLabel;
        private readonly Label _xmlPathLabel;
        private readonly FlowLayoutPanel _resultsFlow;
        private readonly List<Control> _resultCards = new List<Control>();
        private const int AutoCollapseSectionThreshold = 120;
        private const int MaxExpandableBodyHeight = 1400;

        internal ReportViewerForm(AuditReportData data)
        {
            _data = data;
            _failedByCheck = data.ModelCheckerFailedElements
                .GroupBy(v => BuildCheckKey(v.Heading, v.Section, v.CheckName), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

            Text = "Herramienta de auditor\u00EDa JeiAudit | Reporte | Desarrollado por Jason Rojas Estrada - Coordinador BIM";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 1700;
            Height = 980;
            MinimumSize = new Size(1200, 760);
            BackColor = Color.FromArgb(236, 236, 236);
            AutoScaleMode = AutoScaleMode.Dpi;

            Panel header = BuildHeader();
            Panel footer = BuildFooter();

            Panel content = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20, 16, 20, 16),
                BackColor = Color.FromArgb(236, 236, 236)
            };
            Controls.Add(content);
            Controls.Add(footer);
            Controls.Add(header);

            Panel resultsCard = BuildCard(content);
            resultsCard.Dock = DockStyle.Fill;
            resultsCard.Padding = new Padding(10, 12, 10, 12);
            _resultsFlow = new FlowLayoutPanel
            {
                Parent = resultsCard,
                Dock = DockStyle.Fill,
                AutoScroll = true,
                WrapContents = false,
                FlowDirection = FlowDirection.TopDown,
                BackColor = Color.White
            };
            _resultsFlow.SizeChanged += (_, _) => ResizeResultCards();

            Panel summaryCard = BuildCard(content);
            summaryCard.Dock = DockStyle.Top;
            summaryCard.Height = 214;
            BuildSummaryCard(summaryCard, out _modelTagLabel, out _scoreLabel, out _summaryLabel, out _reportDateLabel, out _modelPathLabel, out _xmlPathLabel);

            Panel metadataCard = BuildCard(content);
            metadataCard.Dock = DockStyle.Top;
            metadataCard.Height = 172;
            BuildMetadataCard(metadataCard, out _titleValueLabel, out _dateValueLabel, out _authorValueLabel, out _descriptionValueLabel);

            BindData();
        }

        private static Panel BuildHeader()
        {
            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 78,
                BackColor = Color.FromArgb(57, 67, 82)
            };

            var logo = new Panel
            {
                Parent = header,
                Width = 22,
                Height = 22,
                Left = 16,
                Top = 13,
                BackColor = Color.FromArgb(226, 202, 122)
            };
            logo.Paint += (_, e) =>
            {
                using (var pen = new Pen(Color.FromArgb(107, 107, 107), 2))
                {
                    e.Graphics.DrawRectangle(pen, 1, 1, 19, 19);
                }
            };

            header.Controls.Add(new Label
            {
                Text = "HERRAMIENTA DE AUDITOR\u00CDA JEIAUDIT",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 18f, FontStyle.Regular),
                AutoSize = true,
                Left = 54,
                Top = 6
            });
            header.Controls.Add(new Label
            {
                Text = "Desarrollado por Jason Rojas Estrada - Coordinador BIM, Inspirada en herramientas de Autodesk",
                ForeColor = Color.FromArgb(220, 220, 220),
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
                AutoSize = true,
                Left = 54,
                Top = 42
            });

            var help = new Label
            {
                Parent = header,
                Text = "?",
                Width = 30,
                Height = 30,
                Top = 10,
                Left = header.Width - 42,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.FromArgb(57, 67, 82),
                BackColor = Color.FromArgb(208, 213, 220),
                Font = new Font("Segoe UI", 13f, FontStyle.Bold)
            };

            return header;
        }

        private Panel BuildFooter()
        {
            var footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 66,
                BackColor = Color.FromArgb(98, 98, 98)
            };

            var table = new TableLayoutPanel
            {
                Parent = footer,
                Dock = DockStyle.Fill,
                ColumnCount = 7
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.2857f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.2857f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.2857f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.2857f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.2857f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.2857f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.2858f));

            table.Controls.Add(BuildFooterButton("Copia", (_, _) => CopySummaryToClipboard()), 0, 0);
            table.Controls.Add(BuildFooterButton("Html", (_, _) => ExportHtml()), 1, 0);
            table.Controls.Add(BuildFooterButton("Word", (_, _) => ExportWord()), 2, 0);
            table.Controls.Add(BuildFooterButton("PDF", (_, _) => ExportPdf()), 3, 0);
            table.Controls.Add(BuildFooterButton("Excel", (_, _) => OpenExcelReport()), 4, 0);
            table.Controls.Add(BuildFooterButton("AVT", (_, _) => ExportAvtLikeFile()), 5, 0);
            table.Controls.Add(BuildFooterButton("Cerrar", (_, _) => Close()), 6, 0);
            return footer;
        }

        private static Button BuildFooterButton(string text, EventHandler click)
        {
            var button = new Button
            {
                Text = text,
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(98, 98, 98),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 19f, FontStyle.Regular)
            };
            button.FlatAppearance.BorderSize = 0;
            button.Click += click;
            return button;
        }

        private static Panel BuildCard(Control parent)
        {
            return new Panel
            {
                Parent = parent,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 10, 0, 0)
            };
        }

        private static void BuildMetadataCard(
            Panel card,
            out Label titleValue,
            out Label dateValue,
            out Label authorValue,
            out Label descriptionValue)
        {
            var layout = new TableLayoutPanel
            {
                Parent = card,
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 128f));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            var icon = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                BackColor = Color.White
            };
            icon.Paint += (_, e) => DrawCube(e.Graphics);
            layout.Controls.Add(icon, 0, 0);

            var textPanel = new TableLayoutPanel
            {
                Margin = Padding.Empty,
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 10, 12, 8),
                ColumnCount = 2,
                RowCount = 4
            };
            textPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120f));
            textPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            textPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 32f));
            textPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 32f));
            textPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 32f));
            textPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            layout.Controls.Add(textPanel, 1, 0);

            titleValue = BuildKeyValueRow(textPanel, row: 0, caption: "Titulo", multiline: false, captionFontSize: 15f, valueFontSize: 12f);
            dateValue = BuildKeyValueRow(textPanel, row: 1, caption: "Fecha", multiline: false, captionFontSize: 15f, valueFontSize: 12f);
            authorValue = BuildKeyValueRow(textPanel, row: 2, caption: "Autor", multiline: false, captionFontSize: 15f, valueFontSize: 12f);
            descriptionValue = BuildKeyValueRow(textPanel, row: 3, caption: "Descripcion", multiline: true, captionFontSize: 15f, valueFontSize: 12f);
        }

        private static Label BuildKeyValueRow(
            TableLayoutPanel parent,
            int row,
            string caption,
            bool multiline,
            float captionFontSize,
            float valueFontSize)
        {
            var captionLabel = new Label
            {
                Text = caption,
                Font = new Font("Segoe UI", captionFontSize, FontStyle.Regular),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
                Margin = Padding.Empty
            };
            parent.Controls.Add(captionLabel, 0, row);

            var value = new Label
            {
                Text = "-",
                Font = new Font("Segoe UI", valueFontSize, FontStyle.Regular),
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(80, 80, 80),
                AutoEllipsis = !multiline,
                Margin = Padding.Empty,
                Padding = new Padding(0, multiline ? 2 : 0, 0, 0),
                TextAlign = multiline ? ContentAlignment.TopLeft : ContentAlignment.MiddleLeft
            };
            parent.Controls.Add(value, 1, row);
            return value;
        }

        private static void BuildSummaryCard(
            Panel card,
            out Label modelTagLabel,
            out Label scoreLabel,
            out Label summaryLabel,
            out Label dateLabel,
            out Label modelPathLabel,
            out Label xmlPathLabel)
        {
            var tabPanel = new Panel
            {
                Parent = card,
                Dock = DockStyle.Top,
                Height = 36,
                BackColor = Color.White
            };
            modelTagLabel = new Label
            {
                Parent = tabPanel,
                Text = "(modelo)",
                Left = 0,
                Top = 0,
                Width = 420,
                Height = 36,
                BackColor = Color.FromArgb(241, 241, 241),
                ForeColor = Color.FromArgb(106, 106, 106),
                Font = new Font("Segoe UI", 12f, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 10, 0),
                AutoEllipsis = true
            };

            var body = new TableLayoutPanel
            {
                Parent = card,
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                ColumnCount = 2,
                RowCount = 1,
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220f));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            var scorePanel = new Panel
            {
                Margin = Padding.Empty,
                Dock = DockStyle.Fill,
                Padding = new Padding(4, 6, 4, 8)
            };
            body.Controls.Add(scorePanel, 0, 0);
            scoreLabel = new Label
            {
                Parent = scorePanel,
                Dock = DockStyle.Fill,
                Text = "0%",
                Font = new Font("Segoe UI", 56f, FontStyle.Bold),
                ForeColor = Color.FromArgb(213, 39, 39),
                TextAlign = ContentAlignment.MiddleCenter
            };

            var infoPanel = new TableLayoutPanel
            {
                Margin = Padding.Empty,
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 8, 10, 6),
                ColumnCount = 2,
                RowCount = 4
            };
            infoPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 242f));
            infoPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            infoPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36f));
            infoPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36f));
            infoPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36f));
            infoPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36f));
            body.Controls.Add(infoPanel, 1, 0);

            summaryLabel = BuildKeyValueRow(infoPanel, row: 0, caption: "Resumen de chequeos", multiline: false, captionFontSize: 15f, valueFontSize: 12f);
            dateLabel = BuildKeyValueRow(infoPanel, row: 1, caption: "Fecha del informe", multiline: false, captionFontSize: 15f, valueFontSize: 12f);
            modelPathLabel = BuildKeyValueRow(infoPanel, row: 2, caption: "Revit FilePath", multiline: false, captionFontSize: 15f, valueFontSize: 12f);
            xmlPathLabel = BuildKeyValueRow(infoPanel, row: 3, caption: "Archivo Checkset", multiline: false, captionFontSize: 15f, valueFontSize: 12f);
        }

        private void BindData()
        {
            _titleValueLabel.Text = Fallback(_data.ChecksetTitle, _data.ModelTitle, "-");
            _dateValueLabel.Text = Fallback(_data.ChecksetDate, "-");
            _authorValueLabel.Text = Fallback(_data.ChecksetAuthor, "-");
            _descriptionValueLabel.Text = Fallback(_data.ChecksetDescription, "-");

            _modelTagLabel.Text = Fallback(
                Path.GetFileName(_data.ModelPath),
                _data.ModelTitle,
                "(modelo)");

            _summaryLabel.Text = BuildCountersText(CalculateCounters(_data.ModelCheckerSummaryRows), includePassRate: true);
            _reportDateLabel.Text = _data.GeneratedAtLocal.HasValue
                ? _data.GeneratedAtLocal.Value.ToString("dddd, dd 'de' MMMM 'de' yyyy - HH:mm:ss", CultureInfo.GetCultureInfo("es-ES"))
                : DateTime.Now.ToString("dddd, dd 'de' MMMM 'de' yyyy - HH:mm:ss", CultureInfo.GetCultureInfo("es-ES"));
            _modelPathLabel.Text = Fallback(_data.ModelPath, "-");
            _xmlPathLabel.Text = Fallback(_data.ModelCheckerXmlPath, "-");

            if (_data.ExecutedChecks == 0)
            {
                _scoreLabel.ForeColor = Color.FromArgb(116, 116, 116);
            }
            else
            {
                _scoreLabel.ForeColor = _data.FailCount > 0
                    ? Color.FromArgb(213, 39, 39)
                    : Color.FromArgb(97, 160, 56);
            }

            _scoreLabel.Text = _data.PassRateText;
            BuildResults();
        }

        private void BuildResults()
        {
            _resultsFlow.SuspendLayout();
            _resultsFlow.Controls.Clear();
            _resultCards.Clear();
            _resultsFlow.AutoScrollPosition = Point.Empty;

            if (_data.ModelCheckerSummaryRows.Count == 0)
            {
                var empty = new Label
                {
                    Text = "No hay resultados de chequeos XML para mostrar.",
                    AutoSize = false,
                    Width = 600,
                    Height = 38,
                    Font = new Font("Segoe UI", 12f, FontStyle.Regular),
                    ForeColor = Color.FromArgb(90, 90, 90),
                    TextAlign = ContentAlignment.MiddleLeft
                };
                _resultsFlow.Controls.Add(empty);
                _resultsFlow.ResumeLayout();
                return;
            }

            List<ParameterFillRateRow> globalFillRows = BuildGlobalNoValueParameterFillRows();
            if (globalFillRows.Count > 0)
            {
                Panel chartCard = BuildParameterFillChartCard(globalFillRows);
                AddResultCard(chartCard);
            }

            foreach (IGrouping<string, ModelCheckerSummaryItem> headingGroup in _data.ModelCheckerSummaryRows
                .GroupBy(v => Fallback(v.Heading, "(Sin Heading)"), StringComparer.OrdinalIgnoreCase))
            {
                string headingDescription = headingGroup
                    .Select(v => v.HeadingDescription)
                    .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? string.Empty;

                Panel headingCard = BuildExpandableGroupCard(
                    title: headingGroup.Key,
                    subtitle: headingDescription,
                    summary: BuildCountersText(CalculateCounters(headingGroup), includePassRate: true),
                    indent: 0,
                    titleFontSize: 24f,
                    initiallyExpanded: true,
                    bodyPaddingLeft: 12,
                    out FlowLayoutPanel headingBodyFlow);
                AddResultCard(headingCard);

                foreach (IGrouping<string, ModelCheckerSummaryItem> sectionGroup in headingGroup
                    .GroupBy(v => Fallback(v.Section, "(Sin Section)"), StringComparer.OrdinalIgnoreCase))
                {
                    string sectionDescription = sectionGroup
                        .Select(v => v.SectionDescription)
                        .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? string.Empty;

                    Panel sectionCard = BuildExpandableGroupCard(
                        title: sectionGroup.Key,
                        subtitle: sectionDescription,
                        summary: BuildCountersText(CalculateCounters(sectionGroup), includePassRate: true),
                        indent: 0,
                        titleFontSize: 17f,
                        initiallyExpanded: sectionGroup.Count() <= AutoCollapseSectionThreshold,
                        bodyPaddingLeft: 8,
                        out FlowLayoutPanel sectionBodyFlow);
                    AddFlowItem(headingBodyFlow, sectionCard);

                    bool expandFirstFail = true;
                    foreach (ModelCheckerSummaryItem item in sectionGroup)
                    {
                        bool expand = expandFirstFail && string.Equals(item.Status, "FAIL", StringComparison.OrdinalIgnoreCase);
                        if (expand)
                        {
                            expandFirstFail = false;
                        }

                        Panel checkCard = BuildCheckCard(item, expand, 8);
                        AddFlowItem(sectionBodyFlow, checkCard);
                    }
                }
            }

            _resultsFlow.ResumeLayout();
            ResizeResultCards();
        }

        private void AddResultCard(Control card)
        {
            _resultCards.Add(card);
            _resultsFlow.Controls.Add(card);
        }

        private static void AddFlowItem(FlowLayoutPanel flow, Control item)
        {
            flow.Controls.Add(item);
            ResizeFlowChildren(flow);
        }

        private Panel BuildExpandableGroupCard(
            string title,
            string subtitle,
            string summary,
            int indent,
            float titleFontSize,
            bool initiallyExpanded,
            int bodyPaddingLeft,
            out FlowLayoutPanel bodyFlow)
        {
            int headerHeight = string.IsNullOrWhiteSpace(subtitle) ? 50 : 72;
            var panel = new Panel
            {
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(indent, 0, 0, 8),
                Tag = indent,
                Height = headerHeight + 1
            };

            var bodyHost = new Panel
            {
                Parent = panel,
                Dock = DockStyle.Top,
                BackColor = Color.White,
                Padding = new Padding(bodyPaddingLeft, 0, 4, 4),
                Height = 0,
                Visible = initiallyExpanded,
                AutoScroll = true
            };

            FlowLayoutPanel localBodyFlow = new FlowLayoutPanel
            {
                Parent = bodyHost,
                Dock = DockStyle.Top,
                BackColor = Color.White,
                WrapContents = false,
                FlowDirection = FlowDirection.TopDown,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };

            Action refreshGroupLayout = () =>
            {
                if (bodyHost.Visible)
                {
                    int fullBodyHeight = Math.Max(0, localBodyFlow.Height + bodyHost.Padding.Top + bodyHost.Padding.Bottom);
                    int renderedBodyHeight = Math.Min(MaxExpandableBodyHeight, fullBodyHeight);
                    bodyHost.AutoScrollMinSize = fullBodyHeight > renderedBodyHeight
                        ? new Size(0, fullBodyHeight)
                        : Size.Empty;
                    bodyHost.Height = renderedBodyHeight;
                    panel.Height = headerHeight + renderedBodyHeight + 1;
                }
                else
                {
                    bodyHost.AutoScrollMinSize = Size.Empty;
                    bodyHost.Height = 0;
                    panel.Height = headerHeight + 1;
                }
            };

            localBodyFlow.SizeChanged += (_, _) =>
            {
                ResizeFlowChildren(localBodyFlow);
                refreshGroupLayout();
            };
            localBodyFlow.ControlAdded += (_, _) => refreshGroupLayout();
            localBodyFlow.ControlRemoved += (_, _) => refreshGroupLayout();
            bodyFlow = localBodyFlow;

            var header = new Panel
            {
                Parent = panel,
                Dock = DockStyle.Top,
                Height = headerHeight,
                BackColor = Color.White,
                Padding = new Padding(6, 4, 10, 4)
            };

            var toggleButton = new Button
            {
                Parent = header,
                Width = 20,
                Height = 20,
                Left = 2,
                Top = string.IsNullOrWhiteSpace(subtitle) ? 14 : 24,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(118, 118, 118),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Text = initiallyExpanded ? "v" : ">",
                TabStop = false
            };
            toggleButton.FlatAppearance.BorderSize = 0;

            var summaryLabel = new Label
            {
                Parent = header,
                Dock = DockStyle.Right,
                Width = 560,
                Text = summary,
                Font = new Font("Segoe UI", 12f, FontStyle.Regular),
                ForeColor = Color.FromArgb(95, 95, 95),
                TextAlign = ContentAlignment.TopRight,
                AutoEllipsis = true
            };

            var textPanel = new Panel
            {
                Parent = header,
                Dock = DockStyle.Fill,
                Padding = new Padding(24, 0, 0, 0),
                BackColor = Color.White
            };

            textPanel.Controls.Add(new Label
            {
                Parent = textPanel,
                Dock = DockStyle.Top,
                Height = 36,
                Text = title,
                Font = new Font("Segoe UI", titleFontSize, FontStyle.Regular),
                ForeColor = Color.FromArgb(54, 54, 54),
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            });

            if (!string.IsNullOrWhiteSpace(subtitle))
            {
                textPanel.Controls.Add(new Label
                {
                    Parent = textPanel,
                    Dock = DockStyle.Fill,
                    Text = subtitle,
                    Font = new Font("Segoe UI", 11f, FontStyle.Regular),
                    ForeColor = Color.FromArgb(109, 109, 109),
                    TextAlign = ContentAlignment.TopLeft,
                    AutoEllipsis = true
                });
            }

            toggleButton.Click += (_, _) =>
            {
                bodyHost.Visible = !bodyHost.Visible;
                toggleButton.Text = bodyHost.Visible ? "v" : ">";
                refreshGroupLayout();
            };

            refreshGroupLayout();
            return panel;
        }

        private static void ResizeFlowChildren(FlowLayoutPanel flow)
        {
            int baseWidth = Math.Max(220, flow.ClientSize.Width - 8);
            foreach (Control child in flow.Controls)
            {
                int width = Math.Max(220, baseWidth - child.Margin.Left - child.Margin.Right);
                child.Width = width;
                ResizeNestedFlowChildren(child);
            }
        }

        private static void ResizeNestedFlowChildren(Control parent)
        {
            foreach (Control child in parent.Controls)
            {
                if (child is FlowLayoutPanel nestedFlow)
                {
                    ResizeFlowChildren(nestedFlow);
                }
                else
                {
                    ResizeNestedFlowChildren(child);
                }
            }
        }

        private Panel BuildCheckCard(ModelCheckerSummaryItem item, bool expanded, int indent)
        {
            List<ModelCheckerFailedElementItem> failures = GetFailuresFor(item);
            bool hasUsefulReason = !string.IsNullOrWhiteSpace(item.Reason) &&
                !string.Equals(item.Reason.Trim(), "-", StringComparison.Ordinal);
            bool canExpand = failures.Count > 0 || hasUsefulReason;

            var card = new Panel
            {
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(indent, 4, 0, 10),
                Tag = indent
            };

            var detailsPanel = BuildCheckDetailsPanel(item, failures);
            detailsPanel.Parent = card;
            detailsPanel.Dock = DockStyle.Top;
            detailsPanel.Visible = expanded && canExpand;

            var header = new Panel
            {
                Parent = card,
                Dock = DockStyle.Top,
                Height = 82,
                BackColor = Color.White
            };

            var expandButton = new Button
            {
                Parent = header,
                Width = 20,
                Height = 20,
                Left = 8,
                Top = 30,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(118, 118, 118),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Text = detailsPanel.Visible ? "v" : ">",
                TabStop = false,
                Enabled = canExpand
            };
            expandButton.FlatAppearance.BorderSize = 0;

            var statusBadge = BuildStatusBadge(item.Status);
            statusBadge.Parent = header;
            statusBadge.Left = 34;
            statusBadge.Top = 24;

            var titleLabel = new Label
            {
                Parent = header,
                Text = BuildCheckTitle(item),
                Font = new Font("Segoe UI", 16f, FontStyle.Regular),
                AutoSize = false,
                Left = 72,
                Top = 16,
                Height = 30,
                Width = 1280,
                ForeColor = Color.FromArgb(55, 55, 55),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AutoEllipsis = true
            };

            var descriptionLabel = new Label
            {
                Parent = header,
                Text = BuildCheckSubtitle(item),
                Font = new Font("Segoe UI", 11f, FontStyle.Regular),
                AutoSize = false,
                Left = 72,
                Top = 46,
                Height = 30,
                Width = 1280,
                ForeColor = Color.FromArgb(116, 116, 116),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AutoEllipsis = true
            };

            Action refreshHeight = () =>
            {
                card.Height = header.Height + (detailsPanel.Visible ? detailsPanel.Height : 0) + 1;
                expandButton.Text = detailsPanel.Visible ? "v" : ">";
            };

            expandButton.Click += (_, _) =>
            {
                if (!canExpand)
                {
                    return;
                }

                detailsPanel.Visible = !detailsPanel.Visible;
                refreshHeight();
            };

            refreshHeight();
            return card;
        }

        private Panel BuildParameterFillChartCard(List<ParameterFillRateRow> rows)
        {
            int globalCandidates = rows.Sum(v => v.CandidateElements);
            int globalFilled = rows.Sum(v => v.FilledElements);
            double globalFillRate = globalCandidates > 0
                ? globalFilled * 100.0 / globalCandidates
                : 0.0;

            var card = new Panel
            {
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 0, 0, 12),
                Tag = 0,
                Height = 380
            };

            card.Controls.Add(new Label
            {
                Dock = DockStyle.Top,
                Height = 30,
                Text = $"Eje X: parametros | Metrica: % llenado | Global: {globalFillRate:0.#}% ({globalFilled}/{globalCandidates})",
                Font = new Font("Segoe UI", 10f, FontStyle.Regular),
                ForeColor = Color.FromArgb(90, 90, 90),
                Padding = new Padding(12, 0, 0, 0)
            });

            card.Controls.Add(new Label
            {
                Dock = DockStyle.Top,
                Height = 34,
                Text = "KPI Global - Parametros sin valor",
                Font = new Font("Segoe UI", 15f, FontStyle.Regular),
                ForeColor = Color.FromArgb(35, 35, 35),
                Padding = new Padding(12, 4, 0, 0)
            });

            var host = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(10, 0, 10, 10),
                AutoScroll = true
            };
            card.Controls.Add(host);

            var chartCanvas = new Panel
            {
                Parent = host,
                Height = 292,
                BackColor = Color.White
            };

            chartCanvas.Paint += (_, e) => DrawParameterFillChart(e.Graphics, chartCanvas.ClientRectangle, rows);

            Action resizeCanvas = () =>
            {
                int neededWidth = Math.Max(host.ClientSize.Width - 20, 140 + (rows.Count * 94));
                chartCanvas.Width = neededWidth;
            };

            host.SizeChanged += (_, _) =>
            {
                resizeCanvas();
                chartCanvas.Invalidate();
            };

            resizeCanvas();
            return card;
        }

        private static void DrawParameterFillChart(Graphics graphics, Rectangle bounds, List<ParameterFillRateRow> rows)
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.Clear(Color.White);

            if (rows.Count == 0)
            {
                using (var emptyBrush = new SolidBrush(Color.FromArgb(110, 110, 110)))
                using (var emptyFont = new Font("Segoe UI", 11f, FontStyle.Regular))
                {
                    graphics.DrawString("No hay checks de parametros sin valor para graficar.", emptyFont, emptyBrush, 12f, 12f);
                }
                return;
            }

            int left = 56;
            int top = 14;
            int right = 18;
            int bottom = 74;
            int plotWidth = Math.Max(260, bounds.Width - left - right);
            int plotHeight = Math.Max(120, bounds.Height - top - bottom);
            int plotLeft = left;
            int plotTop = top;
            int plotRight = plotLeft + plotWidth;
            int plotBottom = plotTop + plotHeight;

            using (var axisPen = new Pen(Color.FromArgb(152, 152, 152), 1f))
            using (var gridPen = new Pen(Color.FromArgb(228, 232, 238), 1f))
            using (var labelBrush = new SolidBrush(Color.FromArgb(92, 92, 92)))
            using (var smallFont = new Font("Segoe UI", 8.5f, FontStyle.Regular))
            {
                for (int value = 0; value <= 100; value += 20)
                {
                    float ratio = value / 100f;
                    int y = plotBottom - (int)Math.Round(ratio * plotHeight, MidpointRounding.AwayFromZero);
                    graphics.DrawLine(gridPen, plotLeft, y, plotRight, y);
                    graphics.DrawString(value.ToString(CultureInfo.InvariantCulture) + "%", smallFont, labelBrush, 4f, y - 7f);
                }

                graphics.DrawLine(axisPen, plotLeft, plotTop, plotLeft, plotBottom);
                graphics.DrawLine(axisPen, plotLeft, plotBottom, plotRight, plotBottom);
            }

            int slotWidth = Math.Max(42, plotWidth / Math.Max(1, rows.Count));
            int barWidth = Math.Max(20, Math.Min(44, slotWidth - 18));

            using (var valueFont = new Font("Segoe UI", 8.5f, FontStyle.Bold))
            using (var labelFont = new Font("Segoe UI", 8.2f, FontStyle.Regular))
            using (var valueBrush = new SolidBrush(Color.FromArgb(45, 45, 45)))
            using (var labelBrush = new SolidBrush(Color.FromArgb(75, 75, 75)))
            using (var barBorderPen = new Pen(Color.FromArgb(118, 118, 118), 1f))
            {
                for (int i = 0; i < rows.Count; i++)
                {
                    ParameterFillRateRow row = rows[i];
                    double fill = Math.Max(0.0, Math.Min(100.0, row.FillPercent));
                    int x = plotLeft + (i * slotWidth) + ((slotWidth - barWidth) / 2);
                    int barHeight = (int)Math.Round((fill / 100.0) * plotHeight, MidpointRounding.AwayFromZero);
                    int y = plotBottom - barHeight;

                    Color barColor = InterpolateColor(Color.FromArgb(206, 54, 54), Color.FromArgb(96, 162, 56), fill / 100.0);
                    using (var barBrush = new SolidBrush(barColor))
                    {
                        graphics.FillRectangle(barBrush, x, y, barWidth, barHeight);
                    }
                    graphics.DrawRectangle(barBorderPen, x, y, barWidth, barHeight);

                    string valueText = fill.ToString("0.#", CultureInfo.InvariantCulture) + "%";
                    SizeF valueSize = graphics.MeasureString(valueText, valueFont);
                    float valueX = x + (barWidth - valueSize.Width) / 2f;
                    float valueY = Math.Max(0, y - valueSize.Height - 2f);
                    graphics.DrawString(valueText, valueFont, valueBrush, valueX, valueY);

                    string label = BuildCompactParameterLabel(row.ParameterName);
                    RectangleF labelRect = new RectangleF(x - 10f, plotBottom + 6f, barWidth + 20f, 54f);
                    var format = new StringFormat
                    {
                        Alignment = StringAlignment.Center,
                        LineAlignment = StringAlignment.Near,
                        Trimming = StringTrimming.EllipsisCharacter
                    };
                    graphics.DrawString(label, labelFont, labelBrush, labelRect, format);
                }
            }
        }

        private static Color InterpolateColor(Color from, Color to, double t)
        {
            t = Math.Max(0.0, Math.Min(1.0, t));
            int r = from.R + (int)Math.Round((to.R - from.R) * t, MidpointRounding.AwayFromZero);
            int g = from.G + (int)Math.Round((to.G - from.G) * t, MidpointRounding.AwayFromZero);
            int b = from.B + (int)Math.Round((to.B - from.B) * t, MidpointRounding.AwayFromZero);
            return Color.FromArgb(r, g, b);
        }

        private static string BuildCompactParameterLabel(string parameterName)
        {
            string value = Fallback(parameterName, "-").Trim();
            if (value.Length <= 18)
            {
                return value;
            }

            string[] tokens = value.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length >= 2)
            {
                string head = string.Join("_", tokens.Take(Math.Min(2, tokens.Length)));
                string tail = tokens[tokens.Length - 1];
                string compact = head + "_" + tail;
                if (compact.Length <= 18)
                {
                    return compact;
                }
            }

            return value.Substring(0, 17) + "...";
        }

        private static Panel BuildStatusBadge(string status)
        {
            ResolveStatusVisual(status, out Color backColor, out string text);

            var badge = new Panel
            {
                Width = 30,
                Height = 30,
                BackColor = backColor
            };
            badge.Controls.Add(new Label
            {
                Parent = badge,
                Dock = DockStyle.Fill,
                Text = text,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 12f, FontStyle.Bold)
            });
            return badge;
        }

        private Panel BuildCheckDetailsPanel(ModelCheckerSummaryItem item, List<ModelCheckerFailedElementItem> failures)
        {
            List<FamilyNamingReportItem> familyFailures = GetFailedFamiliesFor(item);
            List<SubprojectNamingReportItem> subprojectAlerts = GetAlertSubprojectsFor(item);
            List<ParameterExistenceReportItem> parameterRows = GetParameterRowsFor(item);

            var panel = new Panel
            {
                BackColor = Color.FromArgb(250, 250, 250),
                Padding = new Padding(72, 2, 18, 12)
            };

            int y = 8;

            string reason = Fallback(item.Reason, "-");
            var reasonLabel = new Label
            {
                Parent = panel,
                Text = reason,
                Left = 0,
                Top = y,
                Width = 1200,
                Height = 25,
                Font = new Font("Segoe UI", 11f, FontStyle.Regular),
                ForeColor = Color.FromArgb(87, 87, 87),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AutoEllipsis = true
            };
            y += 28;

            var countLabel = new Label
            {
                Parent = panel,
                Text = $"Contar: {ResolveDetailsCount(item, familyFailures, subprojectAlerts, parameterRows)}",
                Left = 0,
                Top = y,
                Width = 600,
                Height = 25,
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                ForeColor = Color.FromArgb(68, 68, 68)
            };
            y += 30;

            DataGridView? grid = null;
            if (familyFailures.Count > 0)
            {
                grid = BuildFamilyNamingGrid(familyFailures);
            }
            else if (subprojectAlerts.Count > 0)
            {
                grid = BuildSubprojectNamingGrid(subprojectAlerts);
            }
            else if (parameterRows.Count > 0)
            {
                grid = BuildParameterExistenceGrid(parameterRows);
            }
            else if (failures.Count > 0)
            {
                grid = BuildFailureGrid(failures);
            }

            if (grid != null)
            {
                grid.Parent = panel;
                grid.Left = 0;
                grid.Top = y;
                grid.Width = 1200;
                grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                int detailCount = ResolveDetailsCount(item, familyFailures, subprojectAlerts, parameterRows);
                int rowsToShow = Math.Min(9, Math.Max(1, detailCount));
                grid.Height = grid.ColumnHeadersHeight + rowsToShow * 28 + 2;

                y = grid.Bottom + 10;
                panel.Resize += (_, _) => { grid.Width = Math.Max(300, panel.ClientSize.Width - 18); };
            }

            panel.Height = y + 4;
            return panel;
        }

        private static DataGridView BuildFailureGrid(List<ModelCheckerFailedElementItem> rows)
        {
            var grid = new DataGridView
            {
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AllowUserToOrderColumns = false,
                ReadOnly = true,
                MultiSelect = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                GridColor = Color.FromArgb(220, 220, 220),
                AutoGenerateColumns = false
            };

            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(242, 242, 242);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(60, 60, 60);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10f, FontStyle.Regular);
            grid.EnableHeadersVisualStyles = false;
            grid.DefaultCellStyle.Font = new Font("Segoe UI", 10f, FontStyle.Regular);

            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Categoria", Name = "Category", Width = 120 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Tipo", Name = "Type", Width = 280 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Nombre", Name = "Name", Width = 340 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ID. de elemento", Name = "ElementId", Width = 150 });

            foreach (ModelCheckerFailedElementItem row in rows)
            {
                grid.Rows.Add(
                    Fallback(row.Category, "-"),
                    Fallback(row.FamilyOrType, "-"),
                    Fallback(row.ElementName, "-"),
                    row.ElementId.ToString(CultureInfo.InvariantCulture));
            }

            return grid;
        }

        private static DataGridView BuildFamilyNamingGrid(List<FamilyNamingReportItem> rows)
        {
            var grid = CreateBaseGrid();
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Familia", Name = "FamilyName", Width = 290 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Estado", Name = "Status", Width = 90 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Cod. Disc.", Name = "DisciplineCode", Width = 110 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ARG_", Name = "HasArgPrefix", Width = 70 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Nombre", Name = "HasNamePart", Width = 80 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Motivo", Name = "Reason", Width = 540 });

            foreach (FamilyNamingReportItem row in rows)
            {
                grid.Rows.Add(
                    Fallback(row.FamilyName, "-"),
                    Fallback(row.Status, "-"),
                    Fallback(row.DisciplineCode, "-"),
                    Fallback(row.HasArgPrefix, "-"),
                    Fallback(row.HasNamePart, "-"),
                    Fallback(row.Reason, "-"));
            }

            return grid;
        }

        private static DataGridView BuildSubprojectNamingGrid(List<SubprojectNamingReportItem> rows)
        {
            var grid = CreateBaseGrid();
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Subproyecto", Name = "WorksetName", Width = 300 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Estado", Name = "Status", Width = 90 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ARG_", Name = "HasArgPrefix", Width = 70 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Nombre", Name = "HasNameAfterPrefix", Width = 85 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "2do _", Name = "HasSecondSeparator", Width = 80 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Motivo", Name = "Reason", Width = 530 });

            foreach (SubprojectNamingReportItem row in rows)
            {
                grid.Rows.Add(
                    Fallback(row.WorksetName, "-"),
                    Fallback(row.Status, "-"),
                    Fallback(row.HasArgPrefix, "-"),
                    Fallback(row.HasNameAfterPrefix, "-"),
                    Fallback(row.HasSecondSeparator, "-"),
                    Fallback(row.Reason, "-"));
            }

            return grid;
        }

        private static DataGridView BuildParameterExistenceGrid(List<ParameterExistenceReportItem> rows)
        {
            var grid = CreateBaseGrid();
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Parametro", Name = "Parameter", Width = 270 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Categoria", Name = "Category", Width = 140 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Existe exacto", Name = "ExistsByExactName", Width = 95 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Vinculado cat.", Name = "BoundToCategory", Width = 105 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Estado", Name = "Status", Width = 90 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Nombres similares", Name = "SimilarNames", Width = 470 });

            foreach (ParameterExistenceReportItem row in rows)
            {
                grid.Rows.Add(
                    Fallback(row.Parameter, "-"),
                    Fallback(row.Category, "-"),
                    Fallback(row.ExistsByExactName, "-"),
                    Fallback(row.BoundToCategory, "-"),
                    Fallback(row.Status, "-"),
                    Fallback(row.SimilarNames, "-"));
            }

            return grid;
        }

        private static DataGridView CreateBaseGrid()
        {
            var grid = new DataGridView
            {
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AllowUserToOrderColumns = false,
                ReadOnly = true,
                MultiSelect = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                GridColor = Color.FromArgb(220, 220, 220),
                AutoGenerateColumns = false
            };

            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(242, 242, 242);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(60, 60, 60);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10f, FontStyle.Regular);
            grid.EnableHeadersVisualStyles = false;
            grid.DefaultCellStyle.Font = new Font("Segoe UI", 10f, FontStyle.Regular);
            return grid;
        }

        private int ResolveDetailsCount(
            ModelCheckerSummaryItem item,
            List<FamilyNamingReportItem> familyFailures,
            List<SubprojectNamingReportItem> subprojectAlerts,
            List<ParameterExistenceReportItem> parameterRows)
        {
            if (familyFailures.Count > 0)
            {
                return familyFailures.Count;
            }

            if (subprojectAlerts.Count > 0)
            {
                return subprojectAlerts.Count;
            }

            if (parameterRows.Count > 0)
            {
                return parameterRows.Count;
            }

            return item.FailedElements;
        }

        private List<ModelCheckerFailedElementItem> GetFailuresFor(ModelCheckerSummaryItem item)
        {
            string key = BuildCheckKey(item.Heading, item.Section, item.CheckName);
            return _failedByCheck.TryGetValue(key, out List<ModelCheckerFailedElementItem> rows)
                ? rows
                : new List<ModelCheckerFailedElementItem>();
        }

        private List<FamilyNamingReportItem> GetFailedFamiliesFor(ModelCheckerSummaryItem item)
        {
            if (!IsFamilyNamingCheck(item))
            {
                return new List<FamilyNamingReportItem>();
            }

            return _data.FamilyNamingRows
                .Where(r => string.Equals(r.Status, "FAIL", StringComparison.OrdinalIgnoreCase))
                .OrderBy(r => r.FamilyName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private List<SubprojectNamingReportItem> GetAlertSubprojectsFor(ModelCheckerSummaryItem item)
        {
            if (!IsSubprojectNamingCheck(item))
            {
                return new List<SubprojectNamingReportItem>();
            }

            return _data.SubprojectNamingRows
                .Where(r =>
                    string.Equals(r.Status, "FAIL", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(r.Status, "WARN", StringComparison.OrdinalIgnoreCase))
                .OrderBy(r => r.WorksetName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private List<ParameterExistenceReportItem> GetParameterRowsFor(ModelCheckerSummaryItem item)
        {
            if (!IsParameterExistenceCheck(item))
            {
                return new List<ParameterExistenceReportItem>();
            }

            string parameterName = TryExtractParameterFromCheckName(item.CheckName);
            IEnumerable<ParameterExistenceReportItem> source = _data.ParameterExistenceRows;
            if (!string.IsNullOrWhiteSpace(parameterName))
            {
                source = source.Where(r => string.Equals(r.Parameter, parameterName, StringComparison.OrdinalIgnoreCase));
            }

            if (string.Equals(item.Status, "FAIL", StringComparison.OrdinalIgnoreCase))
            {
                List<ParameterExistenceReportItem> failing = source
                    .Where(r => !string.Equals(r.Status, "OK", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(r => r.Parameter, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(r => r.Category, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (failing.Count > 0)
                {
                    return failing;
                }
            }

            return source
                .OrderBy(r => r.Parameter, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.Category, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static bool IsFamilyNamingCheck(ModelCheckerSummaryItem item)
        {
            return string.Equals(item.CheckName, "JeiAudit_Family_Naming", StringComparison.OrdinalIgnoreCase) ||
                   item.CheckName.IndexOf("Family_Naming", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   item.Section.IndexOf("Nomenclatura_Familias", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsSubprojectNamingCheck(ModelCheckerSummaryItem item)
        {
            return string.Equals(item.CheckName, "JeiAudit_Subproject_Naming", StringComparison.OrdinalIgnoreCase) ||
                   item.CheckName.IndexOf("Subproject_Naming", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   item.Section.IndexOf("Nomenclatura_Subproyectos", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsParameterExistenceCheck(ModelCheckerSummaryItem item)
        {
            return item.CheckName.IndexOf("JeiAudit_Parameter_Existence", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   item.Section.IndexOf("Nomenclatura_Parametros", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string TryExtractParameterFromCheckName(string checkName)
        {
            if (string.IsNullOrWhiteSpace(checkName))
            {
                return string.Empty;
            }

            const string prefix = "JeiAudit_Parameter_Existence_";
            if (!checkName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                return string.Empty;
            }

            string suffix = checkName.Substring(prefix.Length)
                .Replace("_", " ")
                .Trim();
            if (suffix.Length == 0)
            {
                return string.Empty;
            }

            return suffix
                .Replace("CODIGO", "CÓDIGO")
                .Replace("DESCRIPCION", "DESCRIPCIÓN");
        }

        private void ResizeResultCards()
        {
            int baseWidth = Math.Max(320, _resultsFlow.ClientSize.Width - 32);
            foreach (Control card in _resultCards)
            {
                int indent = card.Tag is int indentValue ? indentValue : 0;
                card.Width = Math.Max(280, baseWidth - indent);
                ResizeNestedFlowChildren(card);
            }
        }

        private static GroupCounters CalculateCounters(IEnumerable<ModelCheckerSummaryItem> rows)
        {
            var counters = new GroupCounters();
            foreach (ModelCheckerSummaryItem row in rows)
            {
                counters.Total++;
                if (string.Equals(row.Status, "PASS", StringComparison.OrdinalIgnoreCase))
                {
                    counters.Pass++;
                }
                else if (string.Equals(row.Status, "FAIL", StringComparison.OrdinalIgnoreCase))
                {
                    counters.Fail++;
                }
                else if (string.Equals(row.Status, "COUNT", StringComparison.OrdinalIgnoreCase))
                {
                    counters.CountOnly++;
                }
                else if (string.Equals(row.Status, "UNSUPPORTED", StringComparison.OrdinalIgnoreCase))
                {
                    counters.Unsupported++;
                }
                else
                {
                    counters.Skipped++;
                }
            }

            return counters;
        }

        private static string BuildCountersText(GroupCounters counters, bool includePassRate)
        {
            if (counters.Executed > 0 && includePassRate)
            {
                return $"{counters.Total} chequeos, {counters.Pass} ({counters.PassRateRounded}%) Pass, {counters.Fail} FAIL, {counters.NotExecuted} no ejecutado";
            }

            return $"{counters.Total} chequeos, {counters.NotExecuted} no ejecutado";
        }

        private static string BuildCheckTitle(ModelCheckerSummaryItem item)
        {
            string title = Fallback(item.CheckName, "(check)");
            if (!title.Contains("Obligatorio", StringComparison.OrdinalIgnoreCase) &&
                item.CheckDescription.IndexOf("obligatorio", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                title += " - Obligatorio";
            }

            return title;
        }

        private static string BuildCheckSubtitle(ModelCheckerSummaryItem item)
        {
            if (!string.IsNullOrWhiteSpace(item.CheckDescription))
            {
                return item.CheckDescription;
            }

            if (!string.IsNullOrWhiteSpace(item.Reason))
            {
                return item.Reason;
            }

            return "-";
        }

        private static void ResolveStatusVisual(string status, out Color backColor, out string text)
        {
            if (string.Equals(status, "PASS", StringComparison.OrdinalIgnoreCase))
            {
                backColor = Color.FromArgb(132, 189, 71);
                text = "OK";
                return;
            }

            if (string.Equals(status, "FAIL", StringComparison.OrdinalIgnoreCase))
            {
                backColor = Color.FromArgb(219, 36, 36);
                text = "X";
                return;
            }

            if (string.Equals(status, "COUNT", StringComparison.OrdinalIgnoreCase))
            {
                backColor = Color.FromArgb(53, 109, 193);
                text = "#";
                return;
            }

            if (string.Equals(status, "UNSUPPORTED", StringComparison.OrdinalIgnoreCase))
            {
                backColor = Color.FromArgb(188, 125, 31);
                text = "!";
                return;
            }

            backColor = Color.FromArgb(150, 150, 150);
            text = "-";
        }

        private static void DrawCube(Graphics graphics)
        {
            Point[] top = { new Point(18, 46), new Point(46, 18), new Point(112, 18), new Point(84, 46) };
            Point[] left = { new Point(18, 46), new Point(84, 46), new Point(84, 112), new Point(18, 112) };
            Point[] right = { new Point(84, 46), new Point(112, 18), new Point(112, 84), new Point(84, 112) };

            using (var topBrush = new SolidBrush(Color.FromArgb(231, 231, 231)))
            using (var leftBrush = new SolidBrush(Color.FromArgb(226, 202, 122)))
            using (var rightBrush = new SolidBrush(Color.FromArgb(197, 200, 205)))
            using (var pen = new Pen(Color.FromArgb(107, 107, 107), 3))
            {
                graphics.FillPolygon(topBrush, top);
                graphics.FillPolygon(leftBrush, left);
                graphics.FillPolygon(rightBrush, right);
                graphics.DrawPolygon(pen, top);
                graphics.DrawPolygon(pen, left);
                graphics.DrawPolygon(pen, right);
            }
        }

        private static string BuildCheckKey(string heading, string section, string check)
        {
            return (heading ?? string.Empty).Trim() + "\u001f" +
                   (section ?? string.Empty).Trim() + "\u001f" +
                   (check ?? string.Empty).Trim();
        }

        private static string Fallback(params string[] values)
        {
            foreach (string value in values)
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }

            return string.Empty;
        }

        private void OpenExcelReport()
        {
            if (!File.Exists(_data.ExcelPath))
            {
                MessageBox.Show(this, "No se encontro el archivo Excel del reporte.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Process.Start(new ProcessStartInfo(_data.ExcelPath) { UseShellExecute = true });
        }

        private void ExportHtml()
        {
            string htmlPath = Path.Combine(_data.OutputFolder, "JeiAudit_Informe_Profesional.html");
            File.WriteAllText(htmlPath, BuildHtml(), Encoding.UTF8);
            Process.Start(new ProcessStartInfo(htmlPath) { UseShellExecute = true });
        }

        private void ExportWord()
        {
            string wordPath = Path.Combine(_data.OutputFolder, "JeiAudit_Informe_Profesional.docx");

            try
            {
                BuildWord(wordPath);
                Process.Start(new ProcessStartInfo(wordPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"No se pudo exportar el archivo Word.{Environment.NewLine}{ex.Message}", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportPdf()
        {
            string htmlPath = Path.Combine(_data.OutputFolder, "JeiAudit_Informe_Profesional.html");
            string pdfPath = Path.Combine(_data.OutputFolder, "JeiAudit_Informe_Profesional.pdf");
            File.WriteAllText(htmlPath, BuildHtml(), Encoding.UTF8);

            if (TryGeneratePdfFromHtml(htmlPath, pdfPath, out string error))
            {
                Process.Start(new ProcessStartInfo(pdfPath) { UseShellExecute = true });
                return;
            }

            MessageBox.Show(
                this,
                $"No se pudo generar PDF automatico.{Environment.NewLine}{error}{Environment.NewLine}{Environment.NewLine}Se abrira el HTML para impresion a PDF.",
                "JeiAudit",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            Process.Start(new ProcessStartInfo(htmlPath) { UseShellExecute = true });
        }

        private string BuildHtml()
        {
            List<SectionSummaryRow> sectionRows = BuildSectionSummaryRows();
            List<CriticalFindingRow> criticalRows = BuildCriticalFindingRows(20);
            List<ParameterFillRateRow> parameterFillRows = BuildGlobalNoValueParameterFillRows();
            List<ModelCheckerFailedElementItem> failedSample = _data.ModelCheckerFailedElements
                .Take(150)
                .ToList();
            List<string[]> pebReferences = BuildPebReferenceRows();
            List<FamilyNamingReportItem> familyAlerts = _data.FamilyNamingRows
                .Where(v => string.Equals(v.Status, "FAIL", StringComparison.OrdinalIgnoreCase))
                .OrderBy(v => v.FamilyName, StringComparer.OrdinalIgnoreCase)
                .Take(60)
                .ToList();
            List<SubprojectNamingReportItem> subprojectAlerts = _data.SubprojectNamingRows
                .Where(v =>
                    string.Equals(v.Status, "FAIL", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(v.Status, "WARN", StringComparison.OrdinalIgnoreCase))
                .OrderBy(v => v.WorksetName, StringComparer.OrdinalIgnoreCase)
                .Take(60)
                .ToList();
            List<ParameterExistenceReportItem> parameterDetailRows = _data.ParameterExistenceRows
                .OrderBy(v => string.Equals(v.Status, "OK", StringComparison.OrdinalIgnoreCase) ? 1 : 0)
                .ThenBy(v => v.Parameter, StringComparer.OrdinalIgnoreCase)
                .Take(300)
                .ToList();
            int parameterOk = _data.ParameterExistenceRows.Count(v => string.Equals(v.Status, "OK", StringComparison.OrdinalIgnoreCase));
            int parameterMissing = _data.ParameterExistenceRows.Count - parameterOk;

            var sb = new StringBuilder();
            sb.AppendLine("<!doctype html>");
            sb.AppendLine("<html><head><meta charset=\"utf-8\"/>");
            sb.AppendLine("<title>Informe Profesional de Auditoria BIM - JeiAudit</title>");
            sb.AppendLine("<style>");
            sb.AppendLine("@page{size:A4 portrait;margin:10mm}");
            sb.AppendLine("body{font-family:'Segoe UI',Tahoma,Arial,sans-serif;margin:0;background:#eef1f5;color:#1b1f24;line-height:1.45}");
            sb.AppendLine(".wrap{max-width:1120px;margin:18px auto;padding:0 14px}");
            sb.AppendLine(".hero{background:linear-gradient(135deg,#2f435f,#445b79);color:#fff;border-radius:10px;padding:22px 24px;margin-bottom:14px;box-shadow:0 4px 14px rgba(0,0,0,.15)}");
            sb.AppendLine(".hero h1{margin:0;font-size:28px;font-weight:700;letter-spacing:.4px}");
            sb.AppendLine(".hero p{margin:6px 0 0 0;font-size:14px;opacity:.95}");
            sb.AppendLine(".meta{display:grid;grid-template-columns:220px 1fr;gap:4px 14px;margin-top:14px;font-size:13px}");
            sb.AppendLine(".section{background:#fff;border:1px solid #d7dde5;border-radius:10px;padding:16px 18px;margin-bottom:12px;box-shadow:0 2px 10px rgba(0,0,0,.05)}");
            sb.AppendLine(".section h2{margin:0 0 8px 0;color:#1f3b5d;font-size:20px}");
            sb.AppendLine(".lead{font-size:14px;color:#3a4552}");
            sb.AppendLine(".kpis{display:grid;grid-template-columns:repeat(5,minmax(120px,1fr));gap:10px;margin-top:8px}");
            sb.AppendLine(".kpi{border:1px solid #d8dde6;border-radius:8px;background:#f8fafc;padding:10px}");
            sb.AppendLine(".kpi .label{font-size:12px;color:#546273;text-transform:uppercase;letter-spacing:.4px}");
            sb.AppendLine(".kpi .value{font-size:26px;font-weight:700;color:#1f3b5d;line-height:1.1}");
            sb.AppendLine(".kpi.fail .value{color:#b51f2c}.kpi.pass .value{color:#2d7a3f}");
            sb.AppendLine("table{border-collapse:collapse;width:100%;margin-top:8px}");
            sb.AppendLine("th,td{border:1px solid #d7dde5;padding:7px 8px;text-align:left;font-size:12px;vertical-align:top}");
            sb.AppendLine("th{background:#30455f;color:#fff;font-weight:600}");
            sb.AppendLine("tbody tr:nth-child(even){background:#f8fafd}");
            sb.AppendLine(".status-pass{color:#2d7a3f;font-weight:700}.status-fail{color:#b51f2c;font-weight:700}.status-skip{color:#5a6572;font-weight:700}");
            sb.AppendLine(".muted{color:#667385;font-size:12px}");
            sb.AppendLine(".bullet{margin:6px 0 0 0;padding-left:18px}");
            sb.AppendLine(".bullet li{margin:3px 0}");
            sb.AppendLine(".bar-scroll{overflow-x:auto;padding-bottom:4px}");
            sb.AppendLine(".bar-chart{display:flex;align-items:flex-end;gap:10px;min-height:260px;border:1px solid #d7dde5;border-radius:8px;padding:14px 12px 52px 12px;background:#f9fbfd;min-width:760px}");
            sb.AppendLine(".bar-col{width:72px;flex:0 0 72px;text-align:center;position:relative}");
            sb.AppendLine(".bar-wrap{height:200px;display:flex;align-items:flex-end;justify-content:center}");
            sb.AppendLine(".bar{width:34px;border-radius:4px 4px 0 0;border:1px solid rgba(0,0,0,.2)}");
            sb.AppendLine(".bar-value{font-size:11px;font-weight:700;color:#2c3a4a;margin-bottom:4px}");
            sb.AppendLine(".bar-label{font-size:10px;color:#586678;line-height:1.2;max-width:72px;word-wrap:break-word;position:absolute;left:0;right:0;bottom:-44px}");
            sb.AppendLine("@media print{body{background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact}.section{break-inside:auto;page-break-inside:auto}table{break-inside:auto}tr{break-inside:avoid-page;page-break-inside:avoid}}");
            sb.AppendLine("</style></head><body>");

            sb.AppendLine("<div class='wrap'>");
            sb.AppendLine("<div class='hero'>");
            sb.AppendLine("<h1>Informe Profesional de Auditoria BIM</h1>");
            sb.AppendLine("<p>Entrega ejecutiva de control de calidad y cumplimiento del PEB</p>");
            sb.AppendLine("<div class='meta'>");
            sb.AppendLine($"<div><strong>Proyecto</strong></div><div>{EscapeHtml(Fallback(_data.ChecksetTitle, _data.ModelTitle, "-"))}</div>");
            sb.AppendLine($"<div><strong>Modelo auditado</strong></div><div>{EscapeHtml(Fallback(_data.ModelPath, "-"))}</div>");
            sb.AppendLine($"<div><strong>Checkset XML</strong></div><div>{EscapeHtml(Fallback(_data.ModelCheckerXmlPath, "-"))}</div>");
            sb.AppendLine($"<div><strong>Fecha de informe</strong></div><div>{EscapeHtml(Fallback(_reportDateLabel.Text, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)))}</div>");
            sb.AppendLine("</div>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='section'>");
            sb.AppendLine("<h2>1. Resumen Ejecutivo</h2>");
            sb.AppendLine($"<p class='lead'>{EscapeHtml(BuildExecutiveNarrative())}</p>");
            sb.AppendLine("<div class='kpis'>");
            sb.AppendLine($"<div class='kpi'><div class='label'>Chequeos totales</div><div class='value'>{_data.TotalChecks}</div></div>");
            sb.AppendLine($"<div class='kpi pass'><div class='label'>Pass</div><div class='value'>{_data.PassCount}</div></div>");
            sb.AppendLine($"<div class='kpi fail'><div class='label'>Fail</div><div class='value'>{_data.FailCount}</div></div>");
            sb.AppendLine($"<div class='kpi'><div class='label'>No ejecutado</div><div class='value'>{_data.SkippedCount}</div></div>");
            sb.AppendLine($"<div class='kpi'><div class='label'>Cumplimiento</div><div class='value'>{EscapeHtml(_data.PassRateText)}</div></div>");
            sb.AppendLine("</div>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='section'>");
            sb.AppendLine("<h2>2. Alcance y Metodologia</h2>");
            sb.AppendLine("<p class='lead'>La auditoria se ejecuta sobre el modelo Revit activo, evaluando reglas definidas en el checkset XML. El analisis cubre parametros obligatorios, metadatos, nomenclaturas, vistas de coordinacion y controles personalizados de JeiAudit.</p>");
            sb.AppendLine("<ul class='bullet'>");
            sb.AppendLine("<li>Fuente de verdad: archivo checkset XML vigente del proyecto.</li>");
            sb.AppendLine("<li>Criterio de aprobacion: porcentaje de checks en estado PASS sobre checks ejecutados.</li>");
            sb.AppendLine("<li>Trazabilidad: cada incumplimiento conserva seccion, regla y elementos afectados.</li>");
            sb.AppendLine("<li>Alineado al PEB EIMI-AITEC-DS-000-ZZZ-PEB-ZZZ-ZZZ-001 en controles de calidad, convenciones y nomenclaturas.</li>");
            sb.AppendLine("</ul>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='section'>");
            sb.AppendLine("<h2>3. Resultado Consolidado por Seccion</h2>");
            sb.AppendLine("<table><thead><tr>");
            sb.AppendLine("<th>Heading</th><th>Seccion</th><th>Descripcion</th><th>Total</th><th>Pass</th><th>Fail</th><th>No ejecutado</th><th>% Pass</th>");
            sb.AppendLine("</tr></thead><tbody>");
            foreach (SectionSummaryRow row in sectionRows)
            {
                sb.AppendLine("<tr>");
                sb.AppendLine($"<td>{EscapeHtml(row.Heading)}</td>");
                sb.AppendLine($"<td>{EscapeHtml(row.Section)}</td>");
                sb.AppendLine($"<td>{EscapeHtml(row.Description)}</td>");
                sb.AppendLine($"<td>{row.TotalChecks}</td>");
                sb.AppendLine($"<td class='status-pass'>{row.PassChecks}</td>");
                sb.AppendLine($"<td class='status-fail'>{row.FailChecks}</td>");
                sb.AppendLine($"<td>{row.NotExecutedChecks}</td>");
                sb.AppendLine($"<td>{row.PassRate}%</td>");
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</tbody></table>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='section'>");
            sb.AppendLine("<h2>4. Hallazgos Criticos Priorizados</h2>");
            sb.AppendLine("<p class='muted'>Top de reglas con mayor impacto por cantidad de elementos fallidos.</p>");
            sb.AppendLine("<table><thead><tr>");
            sb.AppendLine("<th>Heading</th><th>Seccion</th><th>Regla</th><th>Elementos FAIL</th><th>Candidatos</th><th>Descripcion del hallazgo</th>");
            sb.AppendLine("</tr></thead><tbody>");
            if (criticalRows.Count == 0)
            {
                sb.AppendLine("<tr><td colspan='6'>No se detectaron hallazgos en estado FAIL.</td></tr>");
            }
            else
            {
                foreach (CriticalFindingRow finding in criticalRows)
                {
                    sb.AppendLine("<tr>");
                    sb.AppendLine($"<td>{EscapeHtml(finding.Heading)}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(finding.Section)}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(finding.CheckName)}</td>");
                    sb.AppendLine($"<td class='status-fail'>{finding.FailedElements}</td>");
                    sb.AppendLine($"<td>{finding.CandidateElements}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(finding.Reason)}</td>");
                    sb.AppendLine("</tr>");
                }
            }
            sb.AppendLine("</tbody></table>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='section'>");
            sb.AppendLine("<h2>5. Detalle de Elementos No Conformes (muestra)</h2>");
            sb.AppendLine("<p class='muted'>Se muestra una muestra operativa para trazabilidad inmediata en coordinacion BIM.</p>");
            sb.AppendLine("<table><thead><tr>");
            sb.AppendLine("<th>Heading</th><th>Seccion</th><th>Regla</th><th>ID Elemento</th><th>Categoria</th><th>Tipo/Familia</th><th>Nombre</th>");
            sb.AppendLine("</tr></thead><tbody>");
            if (failedSample.Count == 0)
            {
                sb.AppendLine("<tr><td colspan='7'>No hay elementos fallidos registrados.</td></tr>");
            }
            else
            {
                foreach (ModelCheckerFailedElementItem item in failedSample)
                {
                    sb.AppendLine("<tr>");
                    sb.AppendLine($"<td>{EscapeHtml(item.Heading)}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(item.Section)}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(item.CheckName)}</td>");
                    sb.AppendLine($"<td>{item.ElementId.ToString(CultureInfo.InvariantCulture)}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(item.Category)}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(item.FamilyOrType)}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(item.ElementName)}</td>");
                    sb.AppendLine("</tr>");
                }
            }
            sb.AppendLine("</tbody></table>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='section'>");
            sb.AppendLine("<h2>6. Verificacion de Parametros PEB (detalle)</h2>");
            sb.AppendLine("<p class='lead'>Validacion de existencia exacta y vinculacion de parametros obligatorios PEB en el modelo.</p>");
            sb.AppendLine($"<p class='muted'>Registros evaluados: {_data.ParameterExistenceRows.Count}. OK: <span class='status-pass'>{parameterOk}</span>. MISSING: <span class='status-fail'>{parameterMissing}</span>.</p>");
            if (parameterFillRows.Count > 0)
            {
                int totalCandidates = parameterFillRows.Sum(v => v.CandidateElements);
                int totalFilled = parameterFillRows.Sum(v => v.FilledElements);
                double globalFillRate = totalCandidates > 0 ? totalFilled * 100.0 / totalCandidates : 0.0;

                sb.AppendLine($"<p class='muted'>Checks de parametros sin valor: llenado global <strong>{globalFillRate.ToString("0.#", CultureInfo.InvariantCulture)}%</strong> ({totalFilled}/{totalCandidates}).</p>");
                sb.AppendLine("<div class='bar-scroll'><div class='bar-chart'>");
                foreach (ParameterFillRateRow row in parameterFillRows)
                {
                    double fill = Math.Max(0.0, Math.Min(100.0, row.FillPercent));
                    int red = 206 + (int)Math.Round((96 - 206) * (fill / 100.0), MidpointRounding.AwayFromZero);
                    int green = 54 + (int)Math.Round((162 - 54) * (fill / 100.0), MidpointRounding.AwayFromZero);
                    int blue = 54 + (int)Math.Round((56 - 54) * (fill / 100.0), MidpointRounding.AwayFromZero);
                    string color = $"rgb({red},{green},{blue})";

                    sb.AppendLine("<div class='bar-col'>");
                    sb.AppendLine($"<div class='bar-value'>{fill.ToString("0.#", CultureInfo.InvariantCulture)}%</div>");
                    sb.AppendLine("<div class='bar-wrap'>");
                    sb.AppendLine($"<div class='bar' style='height:{fill.ToString("0.#", CultureInfo.InvariantCulture)}%;background:{color}'></div>");
                    sb.AppendLine("</div>");
                    sb.AppendLine($"<div class='bar-label'>{EscapeHtml(BuildCompactParameterLabel(row.ParameterName))}</div>");
                    sb.AppendLine("</div>");
                }
                sb.AppendLine("</div></div>");
            }
            sb.AppendLine("<table><thead><tr><th>Parametro</th><th>Categoria</th><th>Existe exacto</th><th>Vinculado categoria</th><th>Estado</th><th>Nombres similares</th></tr></thead><tbody>");
            if (parameterDetailRows.Count == 0)
            {
                sb.AppendLine("<tr><td colspan='6'>No hay datos de verificacion de parametros en el reporte cargado.</td></tr>");
            }
            else
            {
                foreach (ParameterExistenceReportItem row in parameterDetailRows)
                {
                    string css = string.Equals(row.Status, "OK", StringComparison.OrdinalIgnoreCase) ? "status-pass" : "status-fail";
                    sb.AppendLine("<tr>");
                    sb.AppendLine($"<td>{EscapeHtml(row.Parameter)}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(row.Category)}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(row.ExistsByExactName)}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(row.BoundToCategory)}</td>");
                    sb.AppendLine($"<td class='{css}'>{EscapeHtml(row.Status)}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(row.SimilarNames)}</td>");
                    sb.AppendLine("</tr>");
                }
            }
            sb.AppendLine("</tbody></table>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='section'>");
            sb.AppendLine("<h2>7. Anexo - Nomenclaturas Especializadas</h2>");
            sb.AppendLine("<p class='lead'>Validaciones adicionales de JeiAudit para control de nomenclatura de familias y subproyectos.</p>");
            sb.AppendLine("<h3>7.1 Familias con incumplimiento</h3>");
            sb.AppendLine("<table><thead><tr><th>Familia</th><th>Estado</th><th>Codigo Disc.</th><th>Motivo</th></tr></thead><tbody>");
            if (familyAlerts.Count == 0)
            {
                sb.AppendLine("<tr><td colspan='4'>Sin incumplimientos de nomenclatura de familias.</td></tr>");
            }
            else
            {
                foreach (FamilyNamingReportItem row in familyAlerts)
                {
                    sb.AppendLine("<tr>");
                    sb.AppendLine($"<td>{EscapeHtml(row.FamilyName)}</td>");
                    sb.AppendLine($"<td class='status-fail'>{EscapeHtml(row.Status)}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(row.DisciplineCode)}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(row.Reason)}</td>");
                    sb.AppendLine("</tr>");
                }
            }
            sb.AppendLine("</tbody></table>");

            sb.AppendLine("<h3>7.2 Subproyectos con alerta</h3>");
            sb.AppendLine("<table><thead><tr><th>Subproyecto</th><th>Estado</th><th>Motivo</th></tr></thead><tbody>");
            if (subprojectAlerts.Count == 0)
            {
                sb.AppendLine("<tr><td colspan='3'>Sin alertas de nomenclatura de subproyectos.</td></tr>");
            }
            else
            {
                foreach (SubprojectNamingReportItem row in subprojectAlerts)
                {
                    string css = string.Equals(row.Status, "FAIL", StringComparison.OrdinalIgnoreCase) ? "status-fail" : "status-skip";
                    sb.AppendLine("<tr>");
                    sb.AppendLine($"<td>{EscapeHtml(row.WorksetName)}</td>");
                    sb.AppendLine($"<td class='{css}'>{EscapeHtml(row.Status)}</td>");
                    sb.AppendLine($"<td>{EscapeHtml(row.Reason)}</td>");
                    sb.AppendLine("</tr>");
                }
            }
            sb.AppendLine("</tbody></table>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='section'>");
            sb.AppendLine("<h2>8. Conclusiones y Recomendaciones</h2>");
            sb.AppendLine("<ul class='bullet'>");
            sb.AppendLine("<li>Priorizar correcciones en reglas con mayor numero de elementos FAIL.</li>");
            sb.AppendLine("<li>Atacar primero hallazgos de datos base para maximizar mejora en indicadores.</li>");
            sb.AppendLine("<li>Reejecutar auditoria despues de correcciones para validar cierre de brechas.</li>");
            sb.AppendLine("</ul>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='section'>");
            sb.AppendLine("<h2>9. Referencias Normativas del PEB</h2>");
            sb.AppendLine("<p class='lead'>Trazabilidad de los controles auditados contra el PEB de referencia. Se incluyen secciones y paginas clave para sustento de la auditoria.</p>");
            sb.AppendLine("<table><thead><tr><th>Seccion PEB</th><th>Pagina</th><th>Punto de control aplicado</th><th>Aplicacion en JeiAudit</th></tr></thead><tbody>");
            foreach (string[] refRow in pebReferences)
            {
                sb.AppendLine("<tr>");
                sb.AppendLine($"<td>{EscapeHtml(refRow[0])}</td>");
                sb.AppendLine($"<td>{EscapeHtml(refRow[1])}</td>");
                sb.AppendLine($"<td>{EscapeHtml(refRow[2])}</td>");
                sb.AppendLine($"<td>{EscapeHtml(refRow[3])}</td>");
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</tbody></table>");
            sb.AppendLine("</div>");

            sb.AppendLine("</div>");
            sb.AppendLine("</body></html>");
            return sb.ToString();
        }

        private void BuildWord(string outputPath)
        {
            List<SectionSummaryRow> sectionRows = BuildSectionSummaryRows();
            List<CriticalFindingRow> criticalRows = BuildCriticalFindingRows(20);
            List<ParameterFillRateRow> parameterFillRows = BuildGlobalNoValueParameterFillRows();
            List<ModelCheckerFailedElementItem> failedSample = _data.ModelCheckerFailedElements
                .Take(150)
                .ToList();
            List<string[]> pebReferences = BuildPebReferenceRows();
            List<FamilyNamingReportItem> familyAlerts = _data.FamilyNamingRows
                .Where(v => string.Equals(v.Status, "FAIL", StringComparison.OrdinalIgnoreCase))
                .OrderBy(v => v.FamilyName, StringComparer.OrdinalIgnoreCase)
                .Take(60)
                .ToList();
            List<SubprojectNamingReportItem> subprojectAlerts = _data.SubprojectNamingRows
                .Where(v =>
                    string.Equals(v.Status, "FAIL", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(v.Status, "WARN", StringComparison.OrdinalIgnoreCase))
                .OrderBy(v => v.WorksetName, StringComparer.OrdinalIgnoreCase)
                .Take(60)
                .ToList();
            List<ParameterExistenceReportItem> parameterDetailRows = _data.ParameterExistenceRows
                .OrderBy(v => string.Equals(v.Status, "OK", StringComparison.OrdinalIgnoreCase) ? 1 : 0)
                .ThenBy(v => v.Parameter, StringComparer.OrdinalIgnoreCase)
                .Take(300)
                .ToList();
            int parameterOk = _data.ParameterExistenceRows.Count(v => string.Equals(v.Status, "OK", StringComparison.OrdinalIgnoreCase));
            int parameterMissing = _data.ParameterExistenceRows.Count - parameterOk;

            using (WordprocessingDocument document = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = document.AddMainDocumentPart();
                mainPart.Document = new W.Document();
                var body = new W.Body();
                mainPart.Document.Append(body);

                body.Append(CreateWordParagraph("INFORME PROFESIONAL DE AUDITORIA BIM", bold: true, sizeHalfPoints: 40));
                body.Append(CreateWordParagraph("Entrega ejecutiva de control de calidad y cumplimiento del PEB", sizeHalfPoints: 24));
                body.Append(CreateWordParagraph($"Proyecto: {Fallback(_data.ChecksetTitle, _data.ModelTitle, "-")}"));
                body.Append(CreateWordParagraph($"Modelo auditado: {Fallback(_data.ModelPath, "-")}"));
                body.Append(CreateWordParagraph($"Checkset XML: {Fallback(_data.ModelCheckerXmlPath, "-")}"));
                body.Append(CreateWordParagraph($"Fecha de informe: {Fallback(_reportDateLabel.Text, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture))}"));
                body.Append(CreateWordParagraph(string.Empty));

                body.Append(CreateWordParagraph("1. Resumen Ejecutivo", bold: true, sizeHalfPoints: 30));
                body.Append(CreateWordParagraph(BuildExecutiveNarrative()));
                body.Append(CreateWordTable(
                    new[] { "Indicador", "Valor" },
                    new[]
                    {
                        new[] { "Chequeos totales", _data.TotalChecks.ToString(CultureInfo.InvariantCulture) },
                        new[] { "Pass", _data.PassCount.ToString(CultureInfo.InvariantCulture) },
                        new[] { "Fail", _data.FailCount.ToString(CultureInfo.InvariantCulture) },
                        new[] { "No ejecutado", _data.SkippedCount.ToString(CultureInfo.InvariantCulture) },
                        new[] { "Porcentaje de cumplimiento", _data.PassRateText }
                    }));
                body.Append(CreateWordParagraph(string.Empty));

                body.Append(CreateWordParagraph("2. Alcance y Metodologia", bold: true, sizeHalfPoints: 30));
                body.Append(CreateWordParagraph("La auditoria se ejecuta sobre el modelo Revit activo, evaluando reglas definidas en el checkset XML y controles personalizados de JeiAudit."));
                body.Append(CreateWordParagraph("- Fuente de verdad: checkset XML vigente del proyecto."));
                body.Append(CreateWordParagraph("- Criterio de aprobacion: porcentaje PASS sobre checks ejecutados."));
                body.Append(CreateWordParagraph("- Trazabilidad: cada incumplimiento conserva seccion, regla y elementos afectados."));
                body.Append(CreateWordParagraph("- Alineado al PEB EIMI-AITEC-DS-000-ZZZ-PEB-ZZZ-ZZZ-001 en controles de calidad, convenciones y nomenclaturas."));
                body.Append(CreateWordParagraph(string.Empty));

                body.Append(CreateWordParagraph("3. Resultado Consolidado por Seccion", bold: true, sizeHalfPoints: 30));
                body.Append(CreateWordTable(
                    new[] { "Heading", "Seccion", "Descripcion", "Total", "Pass", "Fail", "No ejecutado", "% Pass" },
                    sectionRows.Select(v => new[]
                    {
                        v.Heading,
                        v.Section,
                        v.Description,
                        v.TotalChecks.ToString(CultureInfo.InvariantCulture),
                        v.PassChecks.ToString(CultureInfo.InvariantCulture),
                        v.FailChecks.ToString(CultureInfo.InvariantCulture),
                        v.NotExecutedChecks.ToString(CultureInfo.InvariantCulture),
                        v.PassRate.ToString(CultureInfo.InvariantCulture) + "%"
                    })));
                body.Append(CreateWordParagraph(string.Empty));

                body.Append(CreateWordParagraph("4. Hallazgos Criticos Priorizados", bold: true, sizeHalfPoints: 30));
                body.Append(CreateWordParagraph("Top de reglas con mayor impacto por numero de elementos fallidos."));
                if (criticalRows.Count == 0)
                {
                    body.Append(CreateWordParagraph("No se detectaron hallazgos en estado FAIL."));
                }
                else
                {
                    body.Append(CreateWordTable(
                        new[] { "Heading", "Seccion", "Regla", "Elementos FAIL", "Candidatos", "Descripcion del hallazgo" },
                        criticalRows.Select(v => new[]
                        {
                            v.Heading,
                            v.Section,
                            v.CheckName,
                            v.FailedElements.ToString(CultureInfo.InvariantCulture),
                            v.CandidateElements.ToString(CultureInfo.InvariantCulture),
                            v.Reason
                        })));
                }
                body.Append(CreateWordParagraph(string.Empty));

                body.Append(CreateWordParagraph("5. Detalle de Elementos No Conformes (muestra)", bold: true, sizeHalfPoints: 30));
                body.Append(CreateWordParagraph("Muestra operativa para trazabilidad y correccion en coordinacion BIM."));
                if (failedSample.Count == 0)
                {
                    body.Append(CreateWordParagraph("No hay elementos fallidos registrados."));
                }
                else
                {
                    body.Append(CreateWordTable(
                        new[] { "Heading", "Seccion", "Regla", "ID", "Categoria", "Tipo/Familia", "Nombre" },
                        failedSample.Select(v => new[]
                        {
                            v.Heading,
                            v.Section,
                            v.CheckName,
                            v.ElementId.ToString(CultureInfo.InvariantCulture),
                            v.Category,
                            v.FamilyOrType,
                            v.ElementName
                        })));
                }
                body.Append(CreateWordParagraph(string.Empty));

                body.Append(CreateWordParagraph("6. Verificacion de Parametros PEB (detalle)", bold: true, sizeHalfPoints: 30));
                body.Append(CreateWordParagraph($"Registros evaluados: {_data.ParameterExistenceRows.Count}. OK: {parameterOk}. MISSING: {parameterMissing}."));
                if (parameterFillRows.Count > 0)
                {
                    int totalCandidates = parameterFillRows.Sum(v => v.CandidateElements);
                    int totalFilled = parameterFillRows.Sum(v => v.FilledElements);
                    double globalFillRate = totalCandidates > 0 ? totalFilled * 100.0 / totalCandidates : 0.0;
                    body.Append(CreateWordParagraph($"Resumen global (checks de parametros sin valor): {globalFillRate.ToString("0.#", CultureInfo.InvariantCulture)}% de llenado ({totalFilled}/{totalCandidates})."));
                    body.Append(CreateWordTable(
                        new[] { "Parametro", "% Llenado", "Con valor", "Evaluados", "Sin valor" },
                        parameterFillRows.Select(v => new[]
                        {
                            v.ParameterName,
                            v.FillPercent.ToString("0.#", CultureInfo.InvariantCulture) + "%",
                            v.FilledElements.ToString(CultureInfo.InvariantCulture),
                            v.CandidateElements.ToString(CultureInfo.InvariantCulture),
                            v.EmptyElements.ToString(CultureInfo.InvariantCulture)
                        })));
                    body.Append(CreateWordParagraph(string.Empty));
                }

                if (parameterDetailRows.Count == 0)
                {
                    body.Append(CreateWordParagraph("No hay datos de verificacion de parametros en el reporte cargado."));
                }
                else
                {
                    body.Append(CreateWordTable(
                        new[] { "Parametro", "Categoria", "Existe exacto", "Vinculado categoria", "Estado", "Nombres similares" },
                        parameterDetailRows.Select(v => new[]
                        {
                            v.Parameter,
                            v.Category,
                            v.ExistsByExactName,
                            v.BoundToCategory,
                            v.Status,
                            v.SimilarNames
                        })));
                }
                body.Append(CreateWordParagraph(string.Empty));

                body.Append(CreateWordParagraph("7. Anexo - Nomenclaturas Especializadas", bold: true, sizeHalfPoints: 30));
                body.Append(CreateWordParagraph("7.1 Familias con incumplimiento", bold: true, sizeHalfPoints: 24));
                if (familyAlerts.Count == 0)
                {
                    body.Append(CreateWordParagraph("Sin incumplimientos de nomenclatura de familias."));
                }
                else
                {
                    body.Append(CreateWordTable(
                        new[] { "Familia", "Estado", "Codigo Disc.", "Motivo" },
                        familyAlerts.Select(v => new[] { v.FamilyName, v.Status, v.DisciplineCode, v.Reason })));
                }
                body.Append(CreateWordParagraph(string.Empty));

                body.Append(CreateWordParagraph("7.2 Subproyectos con alerta", bold: true, sizeHalfPoints: 24));
                if (subprojectAlerts.Count == 0)
                {
                    body.Append(CreateWordParagraph("Sin alertas de nomenclatura de subproyectos."));
                }
                else
                {
                    body.Append(CreateWordTable(
                        new[] { "Subproyecto", "Estado", "Motivo" },
                        subprojectAlerts.Select(v => new[] { v.WorksetName, v.Status, v.Reason })));
                }
                body.Append(CreateWordParagraph(string.Empty));

                body.Append(CreateWordParagraph("8. Conclusiones y Recomendaciones", bold: true, sizeHalfPoints: 30));
                body.Append(CreateWordParagraph("- Priorizar reglas con mayor cantidad de elementos FAIL."));
                body.Append(CreateWordParagraph("- Corregir primero datos base y nomenclaturas para elevar el indicador global."));
                body.Append(CreateWordParagraph("- Reejecutar auditoria tras ajustes para validar cierre de brechas."));
                body.Append(CreateWordParagraph(string.Empty));

                body.Append(CreateWordParagraph("9. Referencias Normativas del PEB", bold: true, sizeHalfPoints: 30));
                body.Append(CreateWordParagraph("Trazabilidad de los controles auditados contra el PEB de referencia."));
                body.Append(CreateWordTable(
                    new[] { "Seccion PEB", "Pagina", "Punto de control aplicado", "Aplicacion en JeiAudit" },
                    pebReferences.Select(v => new[] { v[0], v[1], v[2], v[3] })));
                mainPart.Document.Save();
            }
        }

        private List<SectionSummaryRow> BuildSectionSummaryRows()
        {
            return _data.ModelCheckerSummaryRows
                .GroupBy(
                    v => new
                    {
                        Heading = Fallback(v.Heading, "(Sin Heading)"),
                        Section = Fallback(v.Section, "(Sin Seccion)"),
                        Description = Fallback(v.SectionDescription, "-")
                    })
                .Select(g =>
                {
                    int pass = g.Count(v => string.Equals(v.Status, "PASS", StringComparison.OrdinalIgnoreCase));
                    int fail = g.Count(v => string.Equals(v.Status, "FAIL", StringComparison.OrdinalIgnoreCase));
                    int total = g.Count();
                    int executed = pass + fail;
                    int passRate = executed > 0
                        ? (int)Math.Round((double)pass * 100.0 / executed, MidpointRounding.AwayFromZero)
                        : 0;
                    return new SectionSummaryRow
                    {
                        Heading = g.Key.Heading,
                        Section = g.Key.Section,
                        Description = g.Key.Description,
                        TotalChecks = total,
                        PassChecks = pass,
                        FailChecks = fail,
                        NotExecutedChecks = total - executed,
                        PassRate = passRate
                    };
                })
                .OrderBy(v => v.Heading, StringComparer.OrdinalIgnoreCase)
                .ThenBy(v => v.Section, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private List<CriticalFindingRow> BuildCriticalFindingRows(int maxItems)
        {
            return _data.ModelCheckerSummaryRows
                .Where(v => string.Equals(v.Status, "FAIL", StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(v => v.FailedElements)
                .ThenByDescending(v => v.CandidateElements)
                .ThenBy(v => v.Heading, StringComparer.OrdinalIgnoreCase)
                .ThenBy(v => v.Section, StringComparer.OrdinalIgnoreCase)
                .ThenBy(v => v.CheckName, StringComparer.OrdinalIgnoreCase)
                .Take(Math.Max(1, maxItems))
                .Select(v => new CriticalFindingRow
                {
                    Heading = Fallback(v.Heading, "-"),
                    Section = Fallback(v.Section, "-"),
                    CheckName = Fallback(v.CheckName, "-"),
                    FailedElements = v.FailedElements,
                    CandidateElements = v.CandidateElements,
                    Reason = Fallback(v.Reason, "-")
                })
                .ToList();
        }

        private List<ParameterFillRateRow> BuildGlobalNoValueParameterFillRows()
        {
            var byParameter = new Dictionary<string, ParameterFillRateRow>(StringComparer.OrdinalIgnoreCase);

            foreach (ModelCheckerSummaryItem item in _data.ModelCheckerSummaryRows)
            {
                if (!IsNoValueParameterCheck(item))
                {
                    continue;
                }

                int candidates = Math.Max(0, item.CandidateElements);
                if (candidates <= 0)
                {
                    continue;
                }

                int empty = Math.Max(0, item.FailedElements);
                if (empty > candidates)
                {
                    empty = candidates;
                }

                string parameterName = ExtractParameterNameFromNoValueCheck(item);
                if (string.IsNullOrWhiteSpace(parameterName))
                {
                    parameterName = "(Parametro)";
                }

                if (!byParameter.TryGetValue(parameterName, out ParameterFillRateRow row))
                {
                    row = new ParameterFillRateRow { ParameterName = parameterName };
                    byParameter[parameterName] = row;
                }

                row.CandidateElements += candidates;
                row.EmptyElements += empty;
            }

            foreach (ParameterFillRateRow row in byParameter.Values)
            {
                if (row.EmptyElements > row.CandidateElements)
                {
                    row.EmptyElements = row.CandidateElements;
                }
            }

            return byParameter.Values
                .OrderBy(v => v.ParameterName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static bool IsNoValueParameterCheck(ModelCheckerSummaryItem item)
        {
            bool isExecuted = string.Equals(item.Status, "PASS", StringComparison.OrdinalIgnoreCase) ||
                              string.Equals(item.Status, "FAIL", StringComparison.OrdinalIgnoreCase);
            if (!isExecuted)
            {
                return false;
            }

            bool nameSaysNoValue = item.CheckName.IndexOf("NoValue", StringComparison.OrdinalIgnoreCase) >= 0;
            bool descSaysNoValue = item.CheckDescription.IndexOf("sin valor", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                   item.CheckDescription.IndexOf("HasNoValue", StringComparison.OrdinalIgnoreCase) >= 0;
            bool referencesParameter = item.CheckName.IndexOf("ARG_", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                       item.CheckDescription.IndexOf("parametro", StringComparison.OrdinalIgnoreCase) >= 0;

            return (nameSaysNoValue || descSaysNoValue) && referencesParameter;
        }

        private static string ExtractParameterNameFromNoValueCheck(ModelCheckerSummaryItem item)
        {
            string fromDescription = ExtractFirstQuotedToken(item.CheckDescription);
            if (!string.IsNullOrWhiteSpace(fromDescription))
            {
                return fromDescription.Trim();
            }

            string checkName = Fallback(item.CheckName, string.Empty);
            int suffixIndex = checkName.IndexOf(" - ", StringComparison.Ordinal);
            if (suffixIndex >= 0)
            {
                checkName = checkName.Substring(0, suffixIndex);
            }

            Match match = Regex.Match(checkName, @"(ARG_[A-Za-z0-9_ÁÉÍÓÚÜÑáéíóúüñ ]+?)_NoValue", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }

            string[] parts = checkName.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            int start = Array.FindIndex(parts, p => p.StartsWith("ARG", StringComparison.OrdinalIgnoreCase));
            int end = Array.FindIndex(parts, p => string.Equals(p, "NoValue", StringComparison.OrdinalIgnoreCase));
            if (start >= 0 && end > start)
            {
                return string.Join("_", parts.Skip(start).Take(end - start));
            }

            return string.Empty;
        }

        private static string ExtractFirstQuotedToken(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            int first = text.IndexOf('\'');
            if (first < 0 || first >= text.Length - 1)
            {
                return string.Empty;
            }

            int second = text.IndexOf('\'', first + 1);
            if (second <= first + 1)
            {
                return string.Empty;
            }

            return text.Substring(first + 1, second - first - 1);
        }

        private string BuildExecutiveNarrative()
        {
            int total = _data.TotalChecks;
            int pass = _data.PassCount;
            int fail = _data.FailCount;
            int notExecuted = _data.SkippedCount + _data.CountOnlyCount + _data.UnsupportedCount;
            return $"Se evaluaron {total} chequeos. El resultado consolidado muestra {pass} en PASS, {fail} en FAIL y {notExecuted} no ejecutados. El cumplimiento efectivo actual es {_data.PassRateText}.";
        }

        private static List<string[]> BuildPebReferenceRows()
        {
            return new List<string[]>
            {
                new[]
                {
                    "4.8 Planes de entrega de informacion (TIDP/MIDP)",
                    "38",
                    "Nomenclatura de archivo para trazabilidad de intercambio de informacion.",
                    "Regla de nomenclatura de archivo (MIDP) auditada en seccion de nombre de modelo."
                },
                new[]
                {
                    "4.10.9 Control de calidad del modelo de informacion",
                    "64",
                    "Control de calidad de datos del modelo y consistencia de informacion.",
                    "Checks de parametros obligatorios y estado de cumplimiento por seccion."
                },
                new[]
                {
                    "4.11.3 Convenciones generales del modelo BIM",
                    "72",
                    "Convenciones de modelado, estandarizacion de informacion y metadatos.",
                    "Validacion de campos requeridos y reglas de nomenclatura vinculadas al PEB."
                },
                new[]
                {
                    "Nomenclatura para Modelo / Plano / Documentos",
                    "66",
                    "Estructura de nombres para archivos, planos y documentos del proyecto.",
                    "Validaciones de nombre de archivo, nombre de vistas/tablas y nombre de planos."
                },
                new[]
                {
                    "Definicion de nomenclatura de familias por especialidad",
                    "67",
                    "Uso de prefijo corporativo y codigos de disciplina para familias BIM.",
                    "Check especializado de nomenclatura de familias con detalle de incumplimientos."
                },
                new[]
                {
                    "Nomenclatura de subproyectos",
                    "68",
                    "Convencion unica para organizacion de subproyectos y modelos vinculados.",
                    "Check especializado de subproyectos/worksets con clasificacion FAIL/WARN."
                },
                new[]
                {
                    "Vista 3D para exportacion a Navisworks",
                    "72",
                    "Exigencia de vistas de coordinacion para procesos de compatibilizacion.",
                    "Auditoria de nombres de vistas, leyendas y tablas para control operativo."
                }
            };
        }

        private bool TryGeneratePdfFromHtml(string htmlPath, string pdfPath, out string error)
        {
            error = string.Empty;
            string[] browserCandidates =
            {
                @"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                @"C:\Program Files\Google\Chrome\Application\chrome.exe",
                @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            };

            string browserPath = browserCandidates.FirstOrDefault(File.Exists);
            if (string.IsNullOrWhiteSpace(browserPath))
            {
                error = "No se encontro un navegador compatible (Edge/Chrome) para exportacion PDF automatica.";
                return false;
            }

            try
            {
                if (File.Exists(pdfPath))
                {
                    File.Delete(pdfPath);
                }

                string uri = new Uri(htmlPath).AbsoluteUri;
                string arguments = $"--headless --disable-gpu --print-to-pdf=\"{pdfPath}\" --print-to-pdf-no-header \"{uri}\"";

                using (var process = new Process())
                {
                    process.StartInfo = new ProcessStartInfo
                    {
                        FileName = browserPath,
                        Arguments = arguments,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardError = true,
                        RedirectStandardOutput = true
                    };

                    process.Start();
                    if (!process.WaitForExit(45000))
                    {
                        try
                        {
                            process.Kill();
                        }
                        catch
                        {
                        }

                        error = "El navegador excedio el tiempo de conversion a PDF.";
                        return false;
                    }

                    if (process.ExitCode != 0)
                    {
                        string stdError = process.StandardError.ReadToEnd();
                        error = string.IsNullOrWhiteSpace(stdError)
                            ? $"Conversor PDF devolvio codigo {process.ExitCode}."
                            : stdError.Trim();
                        return false;
                    }
                }

                if (!File.Exists(pdfPath))
                {
                    error = "No se genero el archivo PDF.";
                    return false;
                }

                var fileInfo = new FileInfo(pdfPath);
                if (fileInfo.Length <= 0)
                {
                    error = "El archivo PDF generado esta vacio.";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private void ExportAvtLikeFile()
        {
            string avtPath = Path.Combine(_data.OutputFolder, "JeiAudit_Report.avt");
            string json = BuildAvtLikeContent();
            File.WriteAllText(avtPath, json, Encoding.UTF8);
            Process.Start(new ProcessStartInfo(avtPath) { UseShellExecute = true });
        }

        private string BuildAvtLikeContent()
        {
            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine("  \"tool\": \"JeiAudit\",");
            sb.AppendLine("  \"format\": \"avt-like\",");
            sb.AppendLine($"  \"generatedAt\": \"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\",");
            sb.AppendLine($"  \"modelPath\": \"{EscapeJson(_data.ModelPath)}\",");
            sb.AppendLine($"  \"xmlPath\": \"{EscapeJson(_data.ModelCheckerXmlPath)}\",");
            sb.AppendLine("  \"summary\": {");
            sb.AppendLine($"    \"total\": {_data.TotalChecks},");
            sb.AppendLine($"    \"pass\": {_data.PassCount},");
            sb.AppendLine($"    \"fail\": {_data.FailCount},");
            sb.AppendLine($"    \"skipped\": {_data.SkippedCount}");
            sb.AppendLine("  },");
            sb.AppendLine("  \"checks\": [");

            for (int i = 0; i < _data.ModelCheckerSummaryRows.Count; i++)
            {
                ModelCheckerSummaryItem item = _data.ModelCheckerSummaryRows[i];
                sb.AppendLine("    {");
                sb.AppendLine($"      \"heading\": \"{EscapeJson(item.Heading)}\",");
                sb.AppendLine($"      \"section\": \"{EscapeJson(item.Section)}\",");
                sb.AppendLine($"      \"checkName\": \"{EscapeJson(item.CheckName)}\",");
                sb.AppendLine($"      \"status\": \"{EscapeJson(item.Status)}\",");
                sb.AppendLine($"      \"candidate\": {item.CandidateElements},");
                sb.AppendLine($"      \"failed\": {item.FailedElements},");
                sb.AppendLine($"      \"reason\": \"{EscapeJson(item.Reason)}\"");
                sb.Append(i == _data.ModelCheckerSummaryRows.Count - 1 ? "    }\n" : "    },\n");
            }

            sb.AppendLine("  ]");
            sb.AppendLine("}");
            return sb.ToString();
        }

        private void CopySummaryToClipboard()
        {
            var lines = new List<string>
            {
                $"PassRate: {_data.PassRateText}",
                $"Total: {_data.TotalChecks}",
                $"Pass: {_data.PassCount}",
                $"Fail: {_data.FailCount}",
                $"Skipped: {_data.SkippedCount}",
                $"CountOnly: {_data.CountOnlyCount}",
                $"Unsupported: {_data.UnsupportedCount}",
                $"ModelPath: {_data.ModelPath}",
                $"XmlPath: {_data.ModelCheckerXmlPath}",
                string.Empty,
                "Checks:"
            };

            foreach (ModelCheckerSummaryItem item in _data.ModelCheckerSummaryRows)
            {
                lines.Add($"{item.Heading} | {item.Section} | {item.CheckName} | {item.Status} | {item.FailedElements}/{item.CandidateElements}");
            }

            Clipboard.SetText(string.Join(Environment.NewLine, lines));
            MessageBox.Show(this, "Resumen copiado al portapapeles.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static W.Paragraph CreateWordParagraph(string text, bool bold = false, int? sizeHalfPoints = null)
        {
            var run = new W.Run();
            var runProperties = new W.RunProperties();

            if (bold)
            {
                runProperties.Append(new W.Bold());
            }

            if (sizeHalfPoints.HasValue)
            {
                runProperties.Append(new W.FontSize { Val = sizeHalfPoints.Value.ToString(CultureInfo.InvariantCulture) });
            }

            if (runProperties.ChildElements.Count > 0)
            {
                run.Append(runProperties);
            }

            run.Append(new W.Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });
            return new W.Paragraph(run);
        }

        private static W.TableCell CreateWordCell(string text, bool header = false)
        {
            var paragraph = new W.Paragraph();
            var run = new W.Run();

            if (header)
            {
                run.Append(new W.RunProperties(new W.Bold(), new W.Color { Val = "FFFFFF" }));
            }

            run.Append(new W.Text(string.IsNullOrWhiteSpace(text) ? "-" : text) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.Append(run);

            var cell = new W.TableCell(paragraph);
            var properties = new W.TableCellProperties(
                new W.TableCellVerticalAlignment { Val = W.TableVerticalAlignmentValues.Center },
                new W.TableCellWidth { Type = W.TableWidthUnitValues.Auto });

            if (header)
            {
                properties.Append(new W.Shading
                {
                    Val = W.ShadingPatternValues.Clear,
                    Fill = "30455F",
                    Color = "auto"
                });
            }

            cell.Append(properties);

            return cell;
        }

        private static W.Table CreateWordTable(IEnumerable<string> headers, IEnumerable<string[]> rows)
        {
            var table = new W.Table();
            table.AppendChild(new W.TableProperties(
                new W.TableBorders(
                    new W.TopBorder { Val = new EnumValue<W.BorderValues>(W.BorderValues.Single), Size = 4, Color = "C8D0DB" },
                    new W.BottomBorder { Val = new EnumValue<W.BorderValues>(W.BorderValues.Single), Size = 4, Color = "C8D0DB" },
                    new W.LeftBorder { Val = new EnumValue<W.BorderValues>(W.BorderValues.Single), Size = 4, Color = "C8D0DB" },
                    new W.RightBorder { Val = new EnumValue<W.BorderValues>(W.BorderValues.Single), Size = 4, Color = "C8D0DB" },
                    new W.InsideHorizontalBorder { Val = new EnumValue<W.BorderValues>(W.BorderValues.Single), Size = 4, Color = "D8DEE7" },
                    new W.InsideVerticalBorder { Val = new EnumValue<W.BorderValues>(W.BorderValues.Single), Size = 4, Color = "D8DEE7" })));

            var headerRow = new W.TableRow();
            foreach (string header in headers)
            {
                headerRow.Append(CreateWordCell(header, header: true));
            }

            table.Append(headerRow);

            foreach (string[] row in rows)
            {
                var tableRow = new W.TableRow();
                foreach (string value in row)
                {
                    tableRow.Append(CreateWordCell(value));
                }

                table.Append(tableRow);
            }

            return table;
        }

        private static string EscapeHtml(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "-";
            }

            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;");
        }

        private static string EscapeJson(string value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            return value
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n");
        }
    }
}
