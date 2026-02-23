using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace JeiAudit
{
    internal sealed class ModelCheckerSetupForm : Form
    {
        private enum NodeKind
        {
            Heading,
            Section
        }

        private sealed class NodeTag
        {
            public NodeTag(NodeKind kind, XElement element, XElement heading)
            {
                Kind = kind;
                Element = element;
                Heading = heading;
            }

            public NodeKind Kind { get; }
            public XElement Element { get; }
            public XElement Heading { get; }
        }

        private Label _fileStatusLabel;
        private Label _filePathLabel;
        private Label _titleValueLabel;
        private Label _dateValueLabel;
        private Label _authorValueLabel;
        private Label _descriptionValueLabel;
        private TreeView _tree;
        private Label _detailsTitleLabel;
        private Label _detailsSubtitleLabel;
        private FlowLayoutPanel _detailsFlow;
        private Button _saveButton = null!;
        private Button _saveCloseButton = null!;
        private Dictionary<XElement, TreeNode> _nodesByElement = new Dictionary<XElement, TreeNode>();

        private XDocument? _xml;
        private bool _suppressTreeEvents;
        private bool _hasChanges;

        internal string SelectedXmlPath { get; private set; } = string.Empty;

        internal ModelCheckerSetupForm(string initialXmlPath)
        {
            Text = "Herramienta de auditor\u00EDa JeiAudit | Configuraci\u00F3n | Desarrollado por Jason Rojas Estrada - Coordinador BIM";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 1040;
            Height = 900;
            MinimumSize = new Size(980, 760);
            BackColor = Color.FromArgb(236, 236, 236);

            Panel header = BuildHeader();
            Panel footer = BuildFooter();
            Panel content = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(24, 14, 24, 24)
            };

            Controls.Add(content);
            Controls.Add(footer);
            Controls.Add(header);

            Panel checksCard = BuildCard(content);
            checksCard.Dock = DockStyle.Fill;
            SplitContainer split = BuildChecksSplit(checksCard);

            _tree = new TreeView
            {
                Parent = split.Panel1,
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                CheckBoxes = true,
                Font = new Font("Segoe UI", 15f, FontStyle.Regular),
                HideSelection = false,
                FullRowSelect = true,
                BackColor = Color.FromArgb(236, 236, 236)
            };
            _tree.AfterSelect += TreeAfterSelect;
            _tree.AfterCheck += TreeAfterCheck;

            Panel rightPanel = new Panel
            {
                Parent = split.Panel2,
                Dock = DockStyle.Fill,
                Padding = new Padding(16, 12, 12, 8)
            };

            Panel detailsListCard = new Panel
            {
                Parent = rightPanel,
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(8, 8, 8, 8)
            };

            Panel detailsHeaderPanel = new Panel
            {
                Parent = rightPanel,
                Dock = DockStyle.Top,
                Height = 78,
                Padding = new Padding(0, 0, 0, 6),
                BackColor = Color.FromArgb(236, 236, 236)
            };

            _detailsSubtitleLabel = new Label
            {
                Parent = detailsHeaderPanel,
                Text = string.Empty,
                Dock = DockStyle.Top,
                Height = 30,
                Font = new Font("Segoe UI", 12f, FontStyle.Regular),
                ForeColor = Color.FromArgb(95, 95, 95),
                AutoEllipsis = true
            };

            _detailsTitleLabel = new Label
            {
                Parent = detailsHeaderPanel,
                Text = "-",
                Dock = DockStyle.Top,
                Height = 42,
                Font = new Font("Segoe UI", 18f, FontStyle.Regular),
                ForeColor = Color.FromArgb(57, 57, 57)
            };
            _detailsFlow = new FlowLayoutPanel
            {
                Parent = detailsListCard,
                Dock = DockStyle.Fill,
                AutoScroll = true,
                WrapContents = false,
                FlowDirection = FlowDirection.TopDown,
                BackColor = Color.White
            };
            _detailsFlow.SizeChanged += (_, _) => ResizeDetailsCards();

            Panel metadataCard = BuildMetadataCard(content);
            _titleValueLabel = BuildMetadataRow(metadataCard, "Titulo", 12);
            _dateValueLabel = BuildMetadataRow(metadataCard, "Fecha", 46);
            _authorValueLabel = BuildMetadataRow(metadataCard, "Autor", 80);
            _descriptionValueLabel = BuildMetadataRow(metadataCard, "Descripcion", 114);

            Panel fileCard = BuildFileCard(content, out _fileStatusLabel, out _filePathLabel, out Panel fileTextPanel);

            var browseButton = new Button
            {
                Parent = fileTextPanel,
                Text = "Seleccionar XML...",
                Width = 174,
                Height = 34,
                Top = 24,
                Left = fileTextPanel.Width - 182,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(110, 110, 110),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            browseButton.FlatAppearance.BorderSize = 0;
            browseButton.Click += (_, _) => BrowseXmlFile();

            _fileStatusLabel.Cursor = Cursors.Hand;
            _fileStatusLabel.Click += (_, _) => BrowseXmlFile();
            _filePathLabel.Cursor = Cursors.Hand;
            _filePathLabel.Click += (_, _) => BrowseXmlFile();

            FormClosing += OnFormClosing;

            if (!string.IsNullOrWhiteSpace(initialXmlPath) && File.Exists(initialXmlPath))
            {
                LoadXml(initialXmlPath);
            }
            else
            {
                ShowNoDetails();
            }

            RefreshButtons();
        }

        private Panel BuildHeader()
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
            return header;
        }

        private Panel BuildFooter()
        {
            var footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                BackColor = Color.FromArgb(98, 98, 98)
            };
            var table = new TableLayoutPanel
            {
                Parent = footer,
                Dock = DockStyle.Fill,
                ColumnCount = 3
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.34f));

            Button cancel = BuildFooterButton("Cancelar");
            cancel.Click += (_, _) => CloseWithPrompt();
            table.Controls.Add(cancel, 0, 0);

            _saveButton = BuildFooterButton("Guardar");
            _saveButton.Click += (_, _) => SaveXml(showMessage: true);
            table.Controls.Add(_saveButton, 1, 0);

            _saveCloseButton = BuildFooterButton("Guardar y cerrar");
            _saveCloseButton.Click += (_, _) =>
            {
                if (SaveXml(showMessage: false))
                {
                    DialogResult = DialogResult.OK;
                    Close();
                }
            };
            table.Controls.Add(_saveCloseButton, 2, 0);

            return footer;
        }

        private static Button BuildFooterButton(string text)
        {
            var button = new Button
            {
                Dock = DockStyle.Fill,
                Text = text,
                Font = new Font("Segoe UI", 23f, FontStyle.Regular),
                BackColor = Color.FromArgb(98, 98, 98),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private static Panel BuildCard(Control parent)
        {
            var panel = new Panel
            {
                Parent = parent,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 10, 0, 0)
            };
            return panel;
        }

        private static SplitContainer BuildChecksSplit(Control parent)
        {
            var split = new SplitContainer
            {
                Parent = parent,
                Dock = DockStyle.Fill,
                SplitterDistance = 380,
                BackColor = Color.FromArgb(236, 236, 236)
            };
            split.Panel1.BackColor = Color.FromArgb(236, 236, 236);
            split.Panel2.BackColor = Color.FromArgb(236, 236, 236);
            return split;
        }

        private static Panel BuildMetadataCard(Control parent)
        {
            Panel card = BuildCard(parent);
            card.Dock = DockStyle.Top;
            card.Height = 160;

            var textPanel = new Panel
            {
                Parent = card,
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 14, 12, 8)
            };

            var iconPanel = new Panel
            {
                Parent = card,
                Width = 120,
                Dock = DockStyle.Left,
                BackColor = Color.White
            };
            iconPanel.Paint += (_, e) => DrawCube(e.Graphics);
            return textPanel;
        }

        private static Panel BuildFileCard(Control parent, out Label fileStatusLabel, out Label filePathLabel, out Panel fileTextPanel)
        {
            Panel card = BuildCard(parent);
            card.Dock = DockStyle.Top;
            card.Height = 86;

            fileTextPanel = new Panel
            {
                Parent = card,
                Dock = DockStyle.Fill,
                Padding = new Padding(8, 10, 190, 8)
            };
            fileTextPanel.Controls.Add(new Label
            {
                Text = "Archivo Checkset actual",
                Font = new Font("Segoe UI", 15f, FontStyle.Regular),
                AutoSize = true,
                Left = 4,
                Top = 4
            });

            filePathLabel = new Label
            {
                Parent = fileTextPanel,
                Text = "-",
                Font = new Font("Segoe UI", 12f, FontStyle.Regular),
                AutoSize = false,
                Left = 4,
                Top = 36,
                Width = 320,
                Height = 26,
                ForeColor = Color.FromArgb(96, 96, 96),
                AutoEllipsis = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            var statePanel = new Panel
            {
                Parent = card,
                Width = 130,
                Dock = DockStyle.Left,
                BackColor = Color.White
            };
            var stateBadge = new Panel
            {
                Parent = statePanel,
                Width = 112,
                Height = 52,
                Left = 16,
                Top = 16,
                BackColor = Color.FromArgb(226, 202, 122)
            };
            stateBadge.Paint += (_, e) =>
            {
                using (var pen = new Pen(Color.FromArgb(107, 107, 107), 2))
                {
                    e.Graphics.DrawRectangle(pen, 16, 15, 24, 18);
                    e.Graphics.DrawLine(pen, 16, 15, 24, 9);
                    e.Graphics.DrawLine(pen, 24, 9, 40, 9);
                }
            };

            fileStatusLabel = new Label
            {
                Parent = stateBadge,
                Text = "Cerrado",
                Font = new Font("Segoe UI", 12f, FontStyle.Regular),
                AutoSize = true,
                Left = 52,
                Top = 15,
                ForeColor = Color.FromArgb(86, 86, 86)
            };
            return card;
        }

        private static void DrawCube(Graphics g)
        {
            var top = new[] { new Point(18, 44), new Point(44, 18), new Point(106, 18), new Point(80, 44) };
            var left = new[] { new Point(18, 44), new Point(80, 44), new Point(80, 106), new Point(18, 106) };
            var right = new[] { new Point(80, 44), new Point(106, 18), new Point(106, 80), new Point(80, 106) };

            using (var topBrush = new SolidBrush(Color.FromArgb(231, 231, 231)))
            using (var leftBrush = new SolidBrush(Color.FromArgb(226, 202, 122)))
            using (var rightBrush = new SolidBrush(Color.FromArgb(197, 200, 205)))
            using (var pen = new Pen(Color.FromArgb(107, 107, 107), 3))
            {
                g.FillPolygon(topBrush, top);
                g.FillPolygon(leftBrush, left);
                g.FillPolygon(rightBrush, right);
                g.DrawPolygon(pen, top);
                g.DrawPolygon(pen, left);
                g.DrawPolygon(pen, right);
            }
        }

        private static Label BuildMetadataRow(Control parent, string caption, int top)
        {
            parent.Controls.Add(new Label
            {
                Text = caption,
                Font = new Font("Segoe UI", 15f, FontStyle.Regular),
                AutoSize = true,
                Left = 8,
                Top = top
            });

            var value = new Label
            {
                Parent = parent,
                Text = "-",
                Font = new Font("Segoe UI", 12f, FontStyle.Regular),
                AutoSize = false,
                Left = 106,
                Top = top + 4,
                Width = 740,
                Height = 24,
                ForeColor = Color.FromArgb(96, 96, 96)
            };
            return value;
        }

        private void OnFormClosing(object? sender, FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK || !_hasChanges)
            {
                return;
            }

            DialogResult confirm = MessageBox.Show(
                this,
                "Hay cambios sin guardar. Deseas cerrar sin guardar?",
                "JeiAudit",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes)
            {
                e.Cancel = true;
            }
        }

        private void BrowseXmlFile()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Abrir un archivo XML checkset";
                dialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
                dialog.CheckFileExists = true;
                dialog.CheckPathExists = true;
                dialog.Multiselect = false;

                if (!string.IsNullOrWhiteSpace(SelectedXmlPath) && File.Exists(SelectedXmlPath))
                {
                    dialog.InitialDirectory = Path.GetDirectoryName(SelectedXmlPath);
                    dialog.FileName = Path.GetFileName(SelectedXmlPath);
                }

                if (dialog.ShowDialog(this) != DialogResult.OK || !File.Exists(dialog.FileName))
                {
                    return;
                }

                LoadXml(dialog.FileName);
            }
        }

        private void LoadXml(string xmlPath)
        {
            _xml = XDocument.Load(xmlPath, LoadOptions.PreserveWhitespace);
            SelectedXmlPath = xmlPath;
            _filePathLabel.Text = xmlPath;
            _fileStatusLabel.Text = "Abierto";
            PopulateMetadata();
            PopulateTree();
            _hasChanges = false;
            RefreshButtons();
            PluginState.SaveLastXmlPath(xmlPath);
        }

        private void PopulateMetadata()
        {
            XElement? root = _xml?.Root;
            _titleValueLabel.Text = Attr(root, "Name", "-");
            _dateValueLabel.Text = Attr(root, "Date", "-dEsconocido-");
            _authorValueLabel.Text = Attr(root, "Author", string.Empty);
            _descriptionValueLabel.Text = Attr(root, "Description", "-");
        }

        private void PopulateTree()
        {
            _suppressTreeEvents = true;
            try
            {
                _nodesByElement.Clear();
                _tree.Nodes.Clear();

                foreach (XElement heading in GetHeadings())
                {
                    var headingNode = new TreeNode(Attr(heading, "HeadingText", "(Sin Heading)"))
                    {
                        Tag = new NodeTag(NodeKind.Heading, heading, heading),
                        Checked = IsHeadingFullyChecked(heading)
                    };
                    _nodesByElement[heading] = headingNode;

                    foreach (XElement section in GetSections(heading))
                    {
                        var sectionNode = new TreeNode(Attr(section, "SectionName", "(Sin Section)"))
                        {
                            Tag = new NodeTag(NodeKind.Section, section, heading),
                            Checked = IsSectionFullyChecked(section)
                        };
                        _nodesByElement[section] = sectionNode;
                        headingNode.Nodes.Add(sectionNode);
                    }

                    _tree.Nodes.Add(headingNode);
                }

                _tree.ExpandAll();
                if (_tree.Nodes.Count > 0)
                {
                    _tree.SelectedNode = _tree.Nodes[0].Nodes.Count > 0 ? _tree.Nodes[0].Nodes[0] : _tree.Nodes[0];
                }
                else
                {
                    ShowNoDetails();
                }
            }
            finally
            {
                _suppressTreeEvents = false;
            }
        }

        private IEnumerable<XElement> GetHeadings()
        {
            if (_xml?.Root == null)
            {
                return Enumerable.Empty<XElement>();
            }

            return _xml.Root.Elements().Where(e => e.Name.LocalName == "Heading");
        }

        private static IEnumerable<XElement> GetSections(XElement heading)
        {
            return heading.Elements().Where(e => e.Name.LocalName == "Section");
        }

        private static IEnumerable<XElement> GetChecks(XElement section)
        {
            return section.Elements().Where(e => e.Name.LocalName == "Check");
        }

        private void TreeAfterSelect(object? sender, TreeViewEventArgs e)
        {
            RenderDetails();
        }

        private void TreeAfterCheck(object? sender, TreeViewEventArgs e)
        {
            if (_suppressTreeEvents || !(e.Node.Tag is NodeTag tag))
            {
                return;
            }

            _suppressTreeEvents = true;
            try
            {
                if (tag.Kind == NodeKind.Heading)
                {
                    SetHeadingChecked(tag.Element, e.Node.Checked);
                }
                else
                {
                    SetSectionChecked(tag.Element, e.Node.Checked);
                    SetChecked(tag.Heading, IsHeadingFullyChecked(tag.Heading));
                }
                SyncNodeChecks();
            }
            finally
            {
                _suppressTreeEvents = false;
            }

            _hasChanges = true;
            RenderDetails();
        }

        private void RenderDetails()
        {
            _detailsFlow.SuspendLayout();
            _detailsFlow.Controls.Clear();

            try
            {
                TreeNode? node = _tree.SelectedNode;
                if (node == null || !(node.Tag is NodeTag tag))
                {
                    ShowNoDetails();
                    return;
                }

                if (tag.Kind == NodeKind.Section)
                {
                    XElement section = tag.Element;
                    _detailsTitleLabel.Text = Attr(section, "SectionName", "(Sin Section)");
                    _detailsSubtitleLabel.Text = Attr(section, "Description", Attr(section, "Title", "-"));
                    RenderChecks(section, includeSectionPrefix: false);
                }
                else
                {
                    XElement heading = tag.Element;
                    _detailsTitleLabel.Text = Attr(heading, "HeadingText", "(Sin Heading)");
                    _detailsSubtitleLabel.Text = Attr(heading, "Description", "-");
                    foreach (XElement section in GetSections(heading))
                    {
                        _detailsFlow.Controls.Add(new Label
                        {
                            Text = Attr(section, "SectionName", "(Sin Section)"),
                            Font = new Font("Segoe UI", 16f, FontStyle.Regular),
                            ForeColor = Color.FromArgb(67, 67, 67),
                            AutoSize = false,
                            Width = GetDetailsCardWidth(),
                            Height = 32,
                            Margin = new Padding(0, 8, 0, 0),
                            BackColor = Color.White
                        });
                        RenderChecks(section, includeSectionPrefix: true);
                    }
                }
            }
            finally
            {
                _detailsFlow.ResumeLayout();
            }
        }

        private void RenderChecks(XElement section, bool includeSectionPrefix)
        {
            string sectionName = Attr(section, "SectionName", "(Sin Section)");
            foreach (XElement check in GetChecks(section))
            {
                string checkName = Attr(check, "CheckName", "(Sin CheckName)");
                if (includeSectionPrefix)
                {
                    checkName = sectionName + " - " + checkName;
                }
                if (IsTrue(Attr(check, "IsRequired", "False")))
                {
                    checkName += " - Obligatorio";
                }

                string description = Attr(check, "Description", "-");
                AddCheckCard(check, checkName, description);
            }
        }

        private void AddCheckCard(XElement check, string title, string description)
        {
            int width = GetDetailsCardWidth();
            var card = new Panel
            {
                Width = width,
                Height = 70,
                BackColor = Color.White,
                Margin = new Padding(0, 0, 0, 6),
                Tag = "CHECK_CARD"
            };

            var checkBox = new CheckBox
            {
                Parent = card,
                Left = 0,
                Top = 2,
                AutoSize = true,
                Font = new Font("Segoe UI", 15f, FontStyle.Regular),
                Text = title,
                Checked = IsTrue(Attr(check, "IsChecked", "True")),
                Tag = check
            };
            checkBox.CheckedChanged += DetailCheckChanged;

            int textWidth = Math.Max(220, width - 52);
            Size textSize = TextRenderer.MeasureText(
                description,
                new Font("Segoe UI", 12f, FontStyle.Regular),
                new Size(textWidth, 1000),
                TextFormatFlags.WordBreak);

            var desc = new Label
            {
                Parent = card,
                Left = 34,
                Top = 34,
                Width = textWidth,
                Height = Math.Max(20, textSize.Height),
                Font = new Font("Segoe UI", 12f, FontStyle.Regular),
                ForeColor = Color.FromArgb(100, 100, 100),
                Text = description
            };
            card.Height = desc.Bottom + 8;
            _detailsFlow.Controls.Add(card);
        }

        private void DetailCheckChanged(object? sender, EventArgs e)
        {
            if (!(sender is CheckBox checkBox) || !(checkBox.Tag is XElement check))
            {
                return;
            }

            SetChecked(check, checkBox.Checked);
            XElement? section = check.Parent?.Name.LocalName == "Section" ? check.Parent : null;
            XElement? heading = section?.Parent?.Name.LocalName == "Heading" ? section.Parent : null;
            if (section != null)
            {
                SetChecked(section, IsSectionFullyChecked(section));
            }
            if (heading != null)
            {
                SetChecked(heading, IsHeadingFullyChecked(heading));
            }

            _suppressTreeEvents = true;
            try
            {
                SyncNodeChecks();
            }
            finally
            {
                _suppressTreeEvents = false;
            }

            _hasChanges = true;
            RefreshButtons();
        }

        private void SetHeadingChecked(XElement heading, bool value)
        {
            SetChecked(heading, value);
            foreach (XElement section in GetSections(heading))
            {
                SetSectionChecked(section, value);
            }
        }

        private void SetSectionChecked(XElement section, bool value)
        {
            SetChecked(section, value);
            foreach (XElement check in GetChecks(section))
            {
                SetChecked(check, value);
            }
        }

        private static bool IsSectionFullyChecked(XElement section)
        {
            List<XElement> checks = GetChecks(section).ToList();
            return checks.Count == 0
                ? IsTrue(Attr(section, "IsChecked", "True"))
                : checks.All(c => IsTrue(Attr(c, "IsChecked", "True")));
        }

        private static bool IsHeadingFullyChecked(XElement heading)
        {
            List<XElement> sections = GetSections(heading).ToList();
            return sections.Count == 0
                ? IsTrue(Attr(heading, "IsChecked", "True"))
                : sections.All(IsSectionFullyChecked);
        }

        private void SyncNodeChecks()
        {
            foreach (KeyValuePair<XElement, TreeNode> pair in _nodesByElement)
            {
                bool desired = pair.Key.Name.LocalName == "Heading"
                    ? IsHeadingFullyChecked(pair.Key)
                    : IsSectionFullyChecked(pair.Key);
                pair.Value.Checked = desired;
            }
        }

        private void ResizeDetailsCards()
        {
            int width = GetDetailsCardWidth();
            foreach (Control control in _detailsFlow.Controls)
            {
                if (control is Panel panel && Equals(panel.Tag, "CHECK_CARD"))
                {
                    panel.Width = width;
                }
                else if (control is Label label)
                {
                    label.Width = width;
                }
            }
        }

        private int GetDetailsCardWidth()
        {
            return Math.Max(240, _detailsFlow.ClientSize.Width - 26);
        }

        private void ShowNoDetails()
        {
            _detailsTitleLabel.Text = "-";
            _detailsSubtitleLabel.Text = "Selecciona un Heading o Section para ver sus checks.";
            _detailsFlow.Controls.Clear();
        }

        private bool SaveXml(bool showMessage)
        {
            if (_xml == null || string.IsNullOrWhiteSpace(SelectedXmlPath))
            {
                MessageBox.Show(this, "Selecciona primero un archivo XML valido.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (!File.Exists(SelectedXmlPath))
            {
                MessageBox.Show(this, "El archivo XML seleccionado no existe.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            _xml.Save(SelectedXmlPath, SaveOptions.None);
            PluginState.SaveLastXmlPath(SelectedXmlPath);
            _hasChanges = false;
            RefreshButtons();

            if (showMessage)
            {
                MessageBox.Show(this, "Configuracion guardada correctamente.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return true;
        }

        private void RefreshButtons()
        {
            bool enabled = _xml != null && !string.IsNullOrWhiteSpace(SelectedXmlPath);
            _saveButton.Enabled = enabled;
            _saveCloseButton.Enabled = enabled;
        }

        private void CloseWithPrompt()
        {
            if (!_hasChanges)
            {
                DialogResult = DialogResult.Cancel;
                Close();
                return;
            }

            DialogResult confirm = MessageBox.Show(
                this,
                "Hay cambios sin guardar. Deseas guardarlos antes de cerrar?",
                "JeiAudit",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (confirm == DialogResult.Cancel)
            {
                return;
            }
            if (confirm == DialogResult.Yes && !SaveXml(showMessage: false))
            {
                return;
            }

            DialogResult = DialogResult.Cancel;
            Close();
        }

        private static void SetChecked(XElement element, bool value)
        {
            element.SetAttributeValue("IsChecked", value ? "True" : "False");
        }

        private static string Attr(XElement? element, string name, string fallback)
        {
            if (element == null)
            {
                return fallback;
            }

            XAttribute? attr = element.Attributes().FirstOrDefault(a => string.Equals(a.Name.LocalName, name, StringComparison.Ordinal));
            string value = (attr?.Value ?? string.Empty).Trim();
            return value.Length == 0 ? fallback : value;
        }

        private static bool IsTrue(string value)
        {
            return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) || value == "1";
        }
    }
}
