using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace JeiAudit
{
    internal sealed class ChecksetEditorForm : Form
    {
        private readonly Label _pathLabel;
        private readonly Label _statusLabel;
        private readonly TreeView _tree;
        private readonly DataGridView _attributesGrid;
        private readonly Label _selectedNodeLabel;

        private readonly TextBox _nameTextBox;
        private readonly TextBox _descriptionTextBox;
        private readonly TextBox _authorTextBox;
        private readonly TextBox _dateTextBox;
        private readonly CheckBox _allowRequiredCheckBox;

        private Button _saveButton;
        private Button _saveAsButton;
        private Button _copyButton;
        private Button _addHeadingButton;
        private Button _addSectionButton;
        private Button _addCheckButton;
        private Button _addFilterButton;
        private Button _duplicateButton;
        private Button _renameButton;
        private Button _moveUpButton;
        private Button _moveDownButton;
        private Button _matrixButton;
        private Button _deleteButton;

        private Label _footerStatusLabel = null!;
        private TextBox _searchTextBox = null!;
        private Button _findNextButton = null!;
        private ContextMenuStrip _treeMenu = null!;
        private int _searchCursor = -1;

        private XDocument? _xml;
        private bool _hasChanges;
        private bool _suppressTreeEvents;
        private bool _suppressGridEvents;
        private bool _suppressMetadataEvents;

        private static readonly string[] BoolValues = { "True", "False" };
        private static readonly string[] FilterCategoryValues =
        {
            "Category",
            "Parameter",
            "APIParameter",
            "TypeOrInstance",
            "Family",
            "Type",
            "Workset",
            "APIType",
            "Level",
            "PhaseCreated",
            "PhaseDemolished",
            "PhaseStatus",
            "DesignOption",
            "View",
            "StructuralType",
            "Host",
            "HostParameter",
            "Room",
            "Space",
            "Redundant",
            "Custom"
        };

        internal string SelectedXmlPath { get; private set; } = string.Empty;

        internal ChecksetEditorForm(string initialXmlPath)
        {
            Text = "Herramienta de auditor\u00EDa JeiAudit | Editor de Checksets | Desarrollado por Jason Rojas Estrada - Coordinador BIM";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 1320;
            Height = 860;
            MinimumSize = new Size(1180, 760);
            BackColor = Color.FromArgb(236, 236, 236);

            Panel header = BuildHeader();
            Panel toolbar = BuildToolbar();
            Panel footer = BuildFooter();
            Panel content = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20, 12, 20, 16)
            };

            Controls.Add(content);
            Controls.Add(footer);
            Controls.Add(toolbar);
            Controls.Add(header);

            Panel mainCard = BuildCard(content);
            mainCard.Dock = DockStyle.Fill;

            Panel fileCard = BuildCard(content);
            fileCard.Dock = DockStyle.Top;
            fileCard.Height = 82;
            _pathLabel = BuildFilePathPanel(fileCard);
            _statusLabel = BuildFileStateBadge(fileCard);

            var split = new SplitContainer
            {
                Parent = mainCard,
                Dock = DockStyle.Fill,
                SplitterDistance = 430,
                BackColor = Color.FromArgb(236, 236, 236)
            };
            split.Panel1.BackColor = Color.FromArgb(236, 236, 236);
            split.Panel2.BackColor = Color.FromArgb(236, 236, 236);

            Panel leftPanel = new Panel
            {
                Parent = split.Panel1,
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 10, 8, 10)
            };

            _tree = new TreeView
            {
                Parent = leftPanel,
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 10.5f, FontStyle.Regular),
                HideSelection = false,
                FullRowSelect = true
            };
            _tree.AfterSelect += TreeAfterSelect;
            leftPanel.Controls.Add(new Label
            {
                Text = "Estructura del checkset",
                Dock = DockStyle.Top,
                Height = 30,
                Font = new Font("Segoe UI", 12f, FontStyle.Bold),
                ForeColor = Color.FromArgb(65, 65, 65)
            });

            Panel rightPanel = new Panel
            {
                Parent = split.Panel2,
                Dock = DockStyle.Fill,
                Padding = new Padding(8, 10, 10, 10)
            };

            Panel attributesCard = BuildCard(rightPanel);
            attributesCard.Dock = DockStyle.Fill;
            attributesCard.Padding = new Padding(12, 8, 12, 12);

            _attributesGrid = new DataGridView
            {
                Parent = attributesCard,
                Dock = DockStyle.Fill,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false,
                BorderStyle = BorderStyle.FixedSingle,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = new Font("Segoe UI", 10f, FontStyle.Regular)
            };
            _attributesGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "AttrName",
                HeaderText = "Atributo",
                FillWeight = 38f
            });
            _attributesGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "AttrValue",
                HeaderText = "Valor",
                FillWeight = 62f
            });
            _attributesGrid.CellValueChanged += AttributesGridChanged;
            _attributesGrid.UserDeletedRow += AttributesGridChanged;
            _attributesGrid.CellBeginEdit += AttributesGridCellBeginEdit;
            _attributesGrid.CellEndEdit += AttributesGridCellEndEdit;
            _attributesGrid.EditingControlShowing += AttributesGridEditingControlShowing;
            _attributesGrid.CurrentCellDirtyStateChanged += AttributesGridCurrentCellDirtyStateChanged;
            _attributesGrid.DataError += AttributesGridDataError;

            _selectedNodeLabel = new Label
            {
                Parent = attributesCard,
                Dock = DockStyle.Top,
                Height = 36,
                Text = "Elemento seleccionado: -",
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                ForeColor = Color.FromArgb(70, 70, 70),
                AutoEllipsis = true
            };

            Panel metadataCard = BuildCard(rightPanel);
            metadataCard.Dock = DockStyle.Top;
            metadataCard.Height = 188;
            metadataCard.Padding = new Padding(12, 10, 12, 10);
            BuildMetadataEditor(metadataCard,
                out _nameTextBox,
                out _descriptionTextBox,
                out _authorTextBox,
                out _dateTextBox,
                out _allowRequiredCheckBox);

            _saveButton = new Button();
            _saveAsButton = new Button();
            _copyButton = new Button();
            _addHeadingButton = new Button();
            _addSectionButton = new Button();
            _addCheckButton = new Button();
            _addFilterButton = new Button();
            _duplicateButton = new Button();
            _renameButton = new Button();
            _moveUpButton = new Button();
            _moveDownButton = new Button();
            _matrixButton = new Button();
            _deleteButton = new Button();
            ConfigureToolbarButtons(toolbar);
            ConfigureTreeContextMenu();

            FormClosing += OnFormClosing;

            if (!string.IsNullOrWhiteSpace(initialXmlPath) && File.Exists(initialXmlPath))
            {
                LoadXml(initialXmlPath);
            }
            else
            {
                CreateNewCheckset();
            }

            RefreshUiState();
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

            return header;
        }

        private static Panel BuildToolbar()
        {
            return new Panel
            {
                Dock = DockStyle.Top,
                Height = 52,
                BackColor = Color.FromArgb(225, 225, 225),
                Padding = new Padding(12, 8, 12, 8)
            };
        }

        private Panel BuildFooter()
        {
            var footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                BackColor = Color.FromArgb(98, 98, 98)
            };

            footer.Controls.Add(new Label
            {
                Text = "Desarrollado por Jason Rojas Estrada - Coordinador BIM, Inspirada en herramientas de Autodesk",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10.5f, FontStyle.Regular),
                AutoSize = true,
                Left = 14,
                Top = 14
            });

            _footerStatusLabel = new Label
            {
                Parent = footer,
                Text = "Listo.",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10f, FontStyle.Regular),
                AutoSize = false,
                Width = 560,
                Height = 22,
                Left = 280,
                Top = 14,
                TextAlign = ContentAlignment.MiddleLeft,
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
            };

            return footer;
        }

        private static Panel BuildCard(Control parent)
        {
            return new Panel
            {
                Parent = parent,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 8, 0, 0)
            };
        }
        private static Label BuildFilePathPanel(Control card)
        {
            var panel = new Panel
            {
                Parent = card,
                Dock = DockStyle.Fill,
                Padding = new Padding(8, 10, 10, 8)
            };

            panel.Controls.Add(new Label
            {
                Text = "Archivo checkset actual",
                Font = new Font("Segoe UI", 12.5f, FontStyle.Bold),
                AutoSize = true,
                Left = 4,
                Top = 2
            });

            var path = new Label
            {
                Parent = panel,
                Text = "-",
                Font = new Font("Segoe UI", 10.5f, FontStyle.Regular),
                AutoSize = false,
                Left = 4,
                Top = 30,
                Width = 900,
                Height = 36,
                ForeColor = Color.FromArgb(86, 86, 86),
                AutoEllipsis = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            return path;
        }

        private static Label BuildFileStateBadge(Control card)
        {
            var statePanel = new Panel
            {
                Parent = card,
                Width = 144,
                Dock = DockStyle.Left,
                BackColor = Color.White
            };

            var badge = new Panel
            {
                Parent = statePanel,
                Width = 114,
                Height = 52,
                Left = 14,
                Top = 14,
                BackColor = Color.FromArgb(226, 202, 122)
            };

            var label = new Label
            {
                Parent = badge,
                Text = "Nuevo",
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                AutoSize = true,
                Left = 24,
                Top = 15,
                ForeColor = Color.FromArgb(72, 72, 72)
            };
            return label;
        }

        private void ConfigureToolbarButtons(Panel toolbar)
        {
            var searchPanel = new Panel
            {
                Parent = toolbar,
                Dock = DockStyle.Right,
                Width = 300,
                Padding = new Padding(6, 4, 0, 4)
            };

            var actionsHost = new Panel
            {
                Parent = toolbar,
                Dock = DockStyle.Fill,
                AutoScroll = true
            };

            var flow = new FlowLayoutPanel
            {
                Parent = actionsHost,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0),
                Padding = new Padding(0),
                Height = 34,
                Location = new Point(0, 0)
            };
            actionsHost.VerticalScroll.Enabled = false;
            actionsHost.HorizontalScroll.Enabled = true;

            Button newButton = BuildToolbarButton("Nuevo");
            newButton.Click += (_, _) => CreateNewChecksetWithPrompt();
            flow.Controls.Add(newButton);

            Button loadButton = BuildToolbarButton("Cargar");
            loadButton.Click += (_, _) => LoadXmlWithDialog();
            flow.Controls.Add(loadButton);

            _copyButton = BuildToolbarButton("Copiar checkset");
            _copyButton.Click += (_, _) => SaveXml(forceSaveAs: true, copyOnly: true);
            flow.Controls.Add(_copyButton);

            _saveButton = BuildToolbarButton("Guardar");
            _saveButton.Click += (_, _) => SaveXml(forceSaveAs: false, copyOnly: false);
            flow.Controls.Add(_saveButton);

            _saveAsButton = BuildToolbarButton("Guardar como");
            _saveAsButton.Click += (_, _) => SaveXml(forceSaveAs: true, copyOnly: false);
            flow.Controls.Add(_saveAsButton);

            flow.Controls.Add(BuildSeparator());

            _addHeadingButton = BuildToolbarButton("+ Heading");
            _addHeadingButton.Click += (_, _) => AddHeading();
            flow.Controls.Add(_addHeadingButton);

            _addSectionButton = BuildToolbarButton("+ Section");
            _addSectionButton.Click += (_, _) => AddSection();
            flow.Controls.Add(_addSectionButton);

            _addCheckButton = BuildToolbarButton("+ Check");
            _addCheckButton.Click += (_, _) => AddCheck();
            flow.Controls.Add(_addCheckButton);

            _addFilterButton = BuildToolbarButton("+ Filter");
            _addFilterButton.Click += (_, _) => AddFilter();
            flow.Controls.Add(_addFilterButton);

            _duplicateButton = BuildToolbarButton("Duplicar");
            _duplicateButton.Click += (_, _) => DuplicateSelectedNode();
            flow.Controls.Add(_duplicateButton);

            _renameButton = BuildToolbarButton("Renombrar");
            _renameButton.Click += (_, _) => RenameSelectedNode();
            flow.Controls.Add(_renameButton);

            _moveUpButton = BuildToolbarButton("Subir");
            _moveUpButton.Click += (_, _) => MoveSelectedNode(-1);
            flow.Controls.Add(_moveUpButton);

            _moveDownButton = BuildToolbarButton("Bajar");
            _moveDownButton.Click += (_, _) => MoveSelectedNode(1);
            flow.Controls.Add(_moveDownButton);

            _matrixButton = BuildToolbarButton("Generar matriz");
            _matrixButton.Click += (_, _) => GenerateChecksMatrix();
            flow.Controls.Add(_matrixButton);

            _deleteButton = BuildToolbarButton("Eliminar");
            _deleteButton.Click += (_, _) => DeleteSelectedNode();
            flow.Controls.Add(_deleteButton);

            _findNextButton = BuildToolbarButton("Buscar");
            _findNextButton.Parent = searchPanel;
            _findNextButton.Width = 88;
            _findNextButton.Dock = DockStyle.Right;
            _findNextButton.Margin = new Padding(0);
            _findNextButton.Click += (_, _) => FindNextInTree();

            _searchTextBox = new TextBox
            {
                Parent = searchPanel,
                Dock = DockStyle.Fill,
                Height = 26,
                Margin = new Padding(0),
                Font = new Font("Segoe UI", 10f, FontStyle.Regular)
            };
            _searchTextBox.TextChanged += (_, _) => _searchCursor = -1;
            _searchTextBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    FindNextInTree();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            };
        }

        private static Control BuildSeparator()
        {
            return new Panel
            {
                Width = 10,
                Height = 28,
                Margin = new Padding(8, 0, 6, 0)
            };
        }

        private static Button BuildToolbarButton(string text)
        {
            var button = new Button
            {
                Text = text,
                Width = 108,
                Height = 34,
                Margin = new Padding(4, 0, 4, 0),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(245, 245, 245),
                ForeColor = Color.FromArgb(55, 55, 55),
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            button.FlatAppearance.BorderColor = Color.FromArgb(184, 184, 184);
            button.FlatAppearance.BorderSize = 1;
            return button;
        }

        private void ConfigureTreeContextMenu()
        {
            _treeMenu = new ContextMenuStrip();

            _treeMenu.Items.Add("Agregar Heading", null, (_, _) => AddHeading());
            _treeMenu.Items.Add("Agregar Section", null, (_, _) => AddSection());
            _treeMenu.Items.Add("Agregar Check", null, (_, _) => AddCheck());
            _treeMenu.Items.Add("Agregar Filter", null, (_, _) => AddFilter());
            _treeMenu.Items.Add(new ToolStripSeparator());
            _treeMenu.Items.Add("Duplicar", null, (_, _) => DuplicateSelectedNode());
            _treeMenu.Items.Add("Renombrar", null, (_, _) => RenameSelectedNode());
            _treeMenu.Items.Add("Subir", null, (_, _) => MoveSelectedNode(-1));
            _treeMenu.Items.Add("Bajar", null, (_, _) => MoveSelectedNode(1));
            _treeMenu.Items.Add(new ToolStripSeparator());
            _treeMenu.Items.Add("Eliminar", null, (_, _) => DeleteSelectedNode());

            _tree.ContextMenuStrip = _treeMenu;
            _tree.NodeMouseClick += (_, e) =>
            {
                _tree.SelectedNode = e.Node;
            };
        }

        private static void BuildMetadataEditor(
            Panel parent,
            out TextBox name,
            out TextBox description,
            out TextBox author,
            out TextBox date,
            out CheckBox allowRequired)
        {
            parent.Controls.Add(new Label
            {
                Text = "Metadatos del checkset",
                Font = new Font("Segoe UI", 12f, FontStyle.Bold),
                AutoSize = true,
                Left = 4,
                Top = 2
            });

            name = BuildMetadataField(parent, "Name", 34, out _);
            description = BuildMetadataField(parent, "Description", 66, out _);
            author = BuildMetadataField(parent, "Author", 98, out _);
            date = BuildMetadataField(parent, "Date", 130, out _);

            allowRequired = new CheckBox
            {
                Parent = parent,
                Text = "AllowRequired",
                AutoSize = true,
                Left = 574,
                Top = 36,
                Font = new Font("Segoe UI", 10f, FontStyle.Regular)
            };
        }

        private static TextBox BuildMetadataField(Control parent, string label, int top, out Label labelControl)
        {
            labelControl = new Label
            {
                Parent = parent,
                Text = label,
                AutoSize = true,
                Left = 8,
                Top = top + 4,
                Font = new Font("Segoe UI", 10f, FontStyle.Regular)
            };
            var text = new TextBox
            {
                Parent = parent,
                Left = 96,
                Top = top,
                Width = 460,
                Height = 26,
                Font = new Font("Segoe UI", 10f, FontStyle.Regular)
            };
            return text;
        }

        private void CreateNewChecksetWithPrompt()
        {
            if (!ConfirmCloseOrDiscardChanges())
            {
                return;
            }

            CreateNewCheckset();
        }

        private void CreateNewCheckset()
        {
            XElement root = new XElement("MCSettings",
                new XAttribute("AllowRequired", "True"),
                new XAttribute("Name", "Nuevo Checkset JeiAudit"),
                new XAttribute("Date", DateTime.Now.ToString("yyyy-MM-dd")),
                new XAttribute("Author", Environment.UserName),
                new XAttribute("Description", "Checkset creado con el editor de JeiAudit"));

            XElement heading = CreateHeadingElement("NUEVO_HEADING");
            XElement section = CreateSectionElement("Nueva_Seccion");
            XElement check = CreateCheckElement("Nuevo_Check");
            check.Add(CreateFilterElement());
            section.Add(check);
            heading.Add(section);
            root.Add(heading);

            _xml = new XDocument(new XDeclaration("1.0", "utf-8", null), root);
            SelectedXmlPath = string.Empty;
            _hasChanges = false;
            PopulateFormFromXml();
            SetStatus("Checkset nuevo creado.");
        }

        private void LoadXmlWithDialog()
        {
            if (!ConfirmCloseOrDiscardChanges())
            {
                return;
            }

            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Cargar checkset XML";
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
            try
            {
                _xml = XDocument.Load(xmlPath, LoadOptions.PreserveWhitespace);
                SelectedXmlPath = xmlPath;
                _hasChanges = false;
                PopulateFormFromXml();
                PluginState.SaveLastXmlPath(xmlPath);
                SetStatus("Checkset cargado.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    $"No se pudo cargar el XML.{Environment.NewLine}{ex.Message}",
                    "JeiAudit",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void PopulateFormFromXml()
        {
            PopulateMetadataFields();
            PopulateTree();
            UpdatePathUi();
            RefreshUiState();
        }
        private void PopulateMetadataFields()
        {
            XElement? root = _xml?.Root;
            _suppressMetadataEvents = true;
            try
            {
                _nameTextBox.Text = Attr(root, "Name", string.Empty);
                _descriptionTextBox.Text = Attr(root, "Description", string.Empty);
                _authorTextBox.Text = Attr(root, "Author", string.Empty);
                _dateTextBox.Text = Attr(root, "Date", string.Empty);
                _allowRequiredCheckBox.Checked = IsTrue(Attr(root, "AllowRequired", "True"));
            }
            finally
            {
                _suppressMetadataEvents = false;
            }

            _nameTextBox.TextChanged -= MetadataChanged;
            _descriptionTextBox.TextChanged -= MetadataChanged;
            _authorTextBox.TextChanged -= MetadataChanged;
            _dateTextBox.TextChanged -= MetadataChanged;
            _allowRequiredCheckBox.CheckedChanged -= MetadataChanged;

            _nameTextBox.TextChanged += MetadataChanged;
            _descriptionTextBox.TextChanged += MetadataChanged;
            _authorTextBox.TextChanged += MetadataChanged;
            _dateTextBox.TextChanged += MetadataChanged;
            _allowRequiredCheckBox.CheckedChanged += MetadataChanged;
        }

        private void PopulateTree()
        {
            _suppressTreeEvents = true;
            _tree.BeginUpdate();
            try
            {
                _tree.Nodes.Clear();

                if (_xml?.Root == null)
                {
                    return;
                }

                TreeNode rootNode = BuildTreeNode(_xml.Root);
                _tree.Nodes.Add(rootNode);

                foreach (XElement heading in _xml.Root.Elements().Where(e => e.Name.LocalName == "Heading"))
                {
                    TreeNode headingNode = BuildTreeNode(heading);
                    rootNode.Nodes.Add(headingNode);

                    foreach (XElement section in heading.Elements().Where(e => e.Name.LocalName == "Section"))
                    {
                        TreeNode sectionNode = BuildTreeNode(section);
                        headingNode.Nodes.Add(sectionNode);

                        foreach (XElement check in section.Elements().Where(e => e.Name.LocalName == "Check"))
                        {
                            TreeNode checkNode = BuildTreeNode(check);
                            sectionNode.Nodes.Add(checkNode);

                            foreach (XElement filter in check.Elements().Where(e => e.Name.LocalName == "Filter"))
                            {
                                checkNode.Nodes.Add(BuildTreeNode(filter));
                            }
                        }
                    }
                }

                rootNode.Expand();
                if (rootNode.Nodes.Count > 0)
                {
                    rootNode.Nodes[0].Expand();
                }

                _tree.SelectedNode = rootNode;
            }
            finally
            {
                _tree.EndUpdate();
                _suppressTreeEvents = false;
            }

            ShowSelectedElement();
        }

        private TreeNode BuildTreeNode(XElement element)
        {
            return new TreeNode(BuildNodeText(element)) { Tag = element };
        }

        private static string BuildNodeText(XElement element)
        {
            string kind = element.Name.LocalName;
            if (kind == "MCSettings")
            {
                string name = Attr(element, "Name", "(Sin Name)");
                return "MCSettings - " + name;
            }

            if (kind == "Heading")
            {
                return Attr(element, "HeadingText", "(Sin HeadingText)");
            }

            if (kind == "Section")
            {
                return Attr(element, "SectionName", "(Sin SectionName)");
            }

            if (kind == "Check")
            {
                return Attr(element, "CheckName", "(Sin CheckName)");
            }

            if (kind == "Filter")
            {
                string category = Attr(element, "Category", "-");
                string property = Attr(element, "Property", "-");
                string condition = Attr(element, "Condition", "-");
                return $"Filter: {category} | {property} | {condition}";
            }

            return kind;
        }

        private void TreeAfterSelect(object? sender, TreeViewEventArgs e)
        {
            if (_suppressTreeEvents)
            {
                return;
            }

            ShowSelectedElement();
        }

        private void ShowSelectedElement()
        {
            XElement? element = GetSelectedElement();
            _selectedNodeLabel.Text = "Elemento seleccionado: "
                + (element == null ? "-" : BuildElementPath(element));
            PopulateAttributesGrid(element);
            RefreshUiState();
        }

        private void PopulateAttributesGrid(XElement? element)
        {
            _suppressGridEvents = true;
            try
            {
                _attributesGrid.Rows.Clear();
                if (element == null)
                {
                    return;
                }

                foreach (XAttribute attr in element.Attributes())
                {
                    _attributesGrid.Rows.Add(attr.Name.LocalName, attr.Value);
                }

                ConfigureGuidedEditors(element);
            }
            finally
            {
                _suppressGridEvents = false;
            }
        }

        private void AttributesGridCellBeginEdit(object? sender, DataGridViewCellCancelEventArgs e)
        {
            if (_suppressGridEvents || e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            XElement? element = GetSelectedElement();
            if (element == null)
            {
                return;
            }

            ConfigureGuidedEditorsForRow(element, e.RowIndex);
        }

        private void AttributesGridCellEndEdit(object? sender, DataGridViewCellEventArgs e)
        {
            if (_suppressGridEvents || e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            XElement? element = GetSelectedElement();
            if (element == null)
            {
                return;
            }

            ConfigureGuidedEditors(element);
        }

        private void AttributesGridEditingControlShowing(object? sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (!(e.Control is ComboBox combo))
            {
                return;
            }

            combo.DropDownStyle = ComboBoxStyle.DropDown;
            combo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            combo.AutoCompleteSource = AutoCompleteSource.ListItems;
            combo.Validating -= ComboEditingControlValidating;
            combo.Validating += ComboEditingControlValidating;
        }

        private void ComboEditingControlValidating(object? sender, CancelEventArgs e)
        {
            if (!(sender is ComboBox combo))
            {
                return;
            }

            string text = (combo.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            bool exists = combo.Items.Cast<object>()
                .Any(item => string.Equals(item?.ToString(), text, StringComparison.OrdinalIgnoreCase));
            if (!exists)
            {
                combo.Items.Add(text);
            }
        }

        private void AttributesGridCurrentCellDirtyStateChanged(object? sender, EventArgs e)
        {
            if (_attributesGrid.IsCurrentCellDirty)
            {
                _attributesGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private static void AttributesGridDataError(object? sender, DataGridViewDataErrorEventArgs e)
        {
            e.ThrowException = false;
        }

        private void ConfigureGuidedEditors(XElement element)
        {
            bool previous = _suppressGridEvents;
            _suppressGridEvents = true;
            try
            {
                for (int i = 0; i < _attributesGrid.Rows.Count; i++)
                {
                    if (_attributesGrid.Rows[i].IsNewRow)
                    {
                        continue;
                    }

                    ConfigureGuidedEditorsForRow(element, i);
                }
            }
            finally
            {
                _suppressGridEvents = previous;
            }
        }

        private void ConfigureGuidedEditorsForRow(XElement element, int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= _attributesGrid.Rows.Count)
            {
                return;
            }

            DataGridViewRow row = _attributesGrid.Rows[rowIndex];
            if (row.IsNewRow)
            {
                return;
            }

            string attrName = (row.Cells[0].Value?.ToString() ?? string.Empty).Trim();
            string attrValue = row.Cells[1].Value?.ToString() ?? string.Empty;

            SetComboCellOptions(rowIndex, 0, GetSuggestedAttributeNames(element), attrName);

            List<string> valueOptions = GetSuggestedAttributeValues(element, attrName);
            if (valueOptions.Count > 0)
            {
                SetComboCellOptions(rowIndex, 1, valueOptions, attrValue);
            }
            else
            {
                EnsureTextCell(rowIndex, 1, attrValue);
            }
        }

        private void SetComboCellOptions(int rowIndex, int columnIndex, IEnumerable<string> options, string currentValue)
        {
            DataGridViewRow row = _attributesGrid.Rows[rowIndex];
            string valueText = currentValue ?? string.Empty;

            var items = options
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Select(v => v.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (items.Count == 0)
            {
                EnsureTextCell(rowIndex, columnIndex, valueText);
                return;
            }

            DataGridViewComboBoxCell? comboCell = row.Cells[columnIndex] as DataGridViewComboBoxCell;
            if (comboCell == null)
            {
                comboCell = new DataGridViewComboBoxCell
                {
                    FlatStyle = FlatStyle.Flat,
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
                };
                row.Cells[columnIndex] = comboCell;
            }

            comboCell.Items.Clear();
            foreach (string item in items)
            {
                comboCell.Items.Add(item);
            }

            if (!string.IsNullOrWhiteSpace(valueText) &&
                !items.Any(item => string.Equals(item, valueText, StringComparison.OrdinalIgnoreCase)))
            {
                comboCell.Items.Add(valueText);
            }

            comboCell.Value = string.IsNullOrWhiteSpace(valueText) ? null : (object)valueText;
        }

        private void EnsureTextCell(int rowIndex, int columnIndex, string valueText)
        {
            DataGridViewCell current = _attributesGrid.Rows[rowIndex].Cells[columnIndex];
            if (current is DataGridViewTextBoxCell)
            {
                current.Value = valueText;
                return;
            }

            _attributesGrid.Rows[rowIndex].Cells[columnIndex] = new DataGridViewTextBoxCell
            {
                Value = valueText
            };
        }

        private List<string> GetSuggestedAttributeNames(XElement element)
        {
            string kind = element.Name.LocalName;
            if (string.Equals(kind, "MCSettings", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string> { "AllowRequired", "Name", "Date", "Author", "Description" };
            }

            if (string.Equals(kind, "Heading", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string> { "ID", "HeadingText", "Description", "IsChecked" };
            }

            if (string.Equals(kind, "Section", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string> { "ID", "SectionName", "Title", "Description", "IsChecked" };
            }

            if (string.Equals(kind, "Check", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string>
                {
                    "ID", "CheckName", "Description", "FailureMessage",
                    "ResultCondition", "CheckType", "IsRequired", "IsChecked"
                };
            }

            if (string.Equals(kind, "Filter", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string>
                {
                    "ID", "Operator", "Category", "Property", "Condition", "Value",
                    "CaseInsensitive", "Unit", "UnitClass", "FieldTitle", "UserDefined", "Validation"
                };
            }

            return element.Attributes().Select(a => a.Name.LocalName).ToList();
        }

        private List<string> GetSuggestedAttributeValues(XElement element, string attributeName)
        {
            if (string.IsNullOrWhiteSpace(attributeName))
            {
                return new List<string>();
            }

            string kind = element.Name.LocalName;
            if (IsBoolAttributeName(kind, attributeName))
            {
                return BoolValues.ToList();
            }

            if (string.Equals(kind, "Check", StringComparison.OrdinalIgnoreCase))
            {
                if (string.Equals(attributeName, "ResultCondition", StringComparison.OrdinalIgnoreCase))
                {
                    return new List<string>
                    {
                        "FailMatchingElements",
                        "FailIfAllElementsFail",
                        "PassMatchingElements"
                    };
                }

                if (string.Equals(attributeName, "CheckType", StringComparison.OrdinalIgnoreCase))
                {
                    return new List<string> { "Custom", "BuiltIn", "Parameter", "Name" };
                }
            }

            if (!string.Equals(kind, "Filter", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string>();
            }

            if (string.Equals(attributeName, "Operator", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string> { "And", "Or" };
            }

            if (string.Equals(attributeName, "Category", StringComparison.OrdinalIgnoreCase))
            {
                return FilterCategoryValues.ToList();
            }

            if (string.Equals(attributeName, "Property", StringComparison.OrdinalIgnoreCase))
            {
                return GetSuggestedFilterProperties();
            }

            if (string.Equals(attributeName, "Condition", StringComparison.OrdinalIgnoreCase))
            {
                return GetSuggestedFilterConditions();
            }

            if (string.Equals(attributeName, "Value", StringComparison.OrdinalIgnoreCase))
            {
                return GetSuggestedFilterValues();
            }

            if (string.Equals(attributeName, "Unit", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(attributeName, "UnitClass", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(attributeName, "Validation", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string> { "None" };
            }

            return new List<string>();
        }

        private bool IsBoolAttributeName(string elementKind, string attributeName)
        {
            if (string.Equals(attributeName, "AllowRequired", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(attributeName, "IsChecked", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(attributeName, "IsRequired", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(attributeName, "CaseInsensitive", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(attributeName, "UserDefined", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        private List<string> GetSuggestedFilterProperties()
        {
            string filterCategory = GetGridAttributeValue("Category");
            if (string.Equals(filterCategory, "Category", StringComparison.OrdinalIgnoreCase))
            {
                List<string> categories = CollectSuggestedCategoryPropertiesFromXml();
                if (categories.Count == 0)
                {
                    categories = new List<string>
                    {
                        "OST_Walls",
                        "OST_Doors",
                        "OST_Windows",
                        "OST_Floors",
                        "OST_Rooms",
                        "OST_StructuralFraming",
                        "OST_StructuralColumns",
                        "OST_PipeCurves",
                        "OST_PipeAccessory",
                        "OST_PipeFitting",
                        "OST_PlumbingFixtures",
                        "OST_MechanicalEquipment"
                    };
                }

                return categories;
            }

            if (string.Equals(filterCategory, "Parameter", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(filterCategory, "APIParameter", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(filterCategory, "HostParameter", StringComparison.OrdinalIgnoreCase))
            {
                return CollectSuggestedParameterNamesFromXml();
            }

            if (string.Equals(filterCategory, "TypeOrInstance", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string> { "Is Element Type" };
            }

            if (string.Equals(filterCategory, "Family", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string> { "Family Name" };
            }

            if (string.Equals(filterCategory, "Type", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string> { "Type Name" };
            }

            if (string.Equals(filterCategory, "Workset", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string> { "Workset Name" };
            }

            if (string.Equals(filterCategory, "Custom", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string> { "ParameterName", "FamilyName", "TypeName", "WorksetName" };
            }

            return new List<string>();
        }

        private List<string> GetSuggestedFilterConditions()
        {
            string filterCategory = GetGridAttributeValue("Category");
            if (string.Equals(filterCategory, "Category", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string> { "Included", "Excluded", "Equal", "NotEqual" };
            }

            if (string.Equals(filterCategory, "TypeOrInstance", StringComparison.OrdinalIgnoreCase))
            {
                return new List<string> { "Equal", "NotEqual", "Included", "Excluded" };
            }

            return new List<string>
            {
                "HasNoValue",
                "HasValue",
                "Exists",
                "DoesNotExist",
                "Equal",
                "NotEqual",
                "Contains",
                "NotContains",
                "StartsWith",
                "NotStartsWith",
                "EndsWith",
                "NotEndsWith",
                "MatchesRegex",
                "NotMatchesRegex",
                "In",
                "NotIn",
                "GreaterThan",
                "GreaterOrEqual",
                "LessThan",
                "LessOrEqual",
                "Duplicated"
            };
        }

        private List<string> GetSuggestedFilterValues()
        {
            string filterCategory = GetGridAttributeValue("Category");
            string condition = GetGridAttributeValue("Condition");

            if (string.Equals(filterCategory, "Category", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(filterCategory, "TypeOrInstance", StringComparison.OrdinalIgnoreCase))
            {
                return BoolValues.ToList();
            }

            if (string.Equals(condition, "HasNoValue", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(condition, "HasValue", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(condition, "Exists", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(condition, "DoesNotExist", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(condition, "IsTrue", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(condition, "IsFalse", StringComparison.OrdinalIgnoreCase))
            {
                return BoolValues.ToList();
            }

            if (string.Equals(filterCategory, "Custom", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(GetGridAttributeValue("Property"), "ParameterName", StringComparison.OrdinalIgnoreCase))
            {
                return CollectSuggestedParameterNamesFromXml();
            }

            return new List<string>();
        }

        private string GetGridAttributeValue(string attributeName)
        {
            foreach (DataGridViewRow row in _attributesGrid.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                string key = (row.Cells[0].Value?.ToString() ?? string.Empty).Trim();
                if (!string.Equals(key, attributeName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return (row.Cells[1].Value?.ToString() ?? string.Empty).Trim();
            }

            return string.Empty;
        }

        private List<string> CollectSuggestedCategoryPropertiesFromXml()
        {
            if (_xml?.Root == null)
            {
                return new List<string>();
            }

            return _xml.Root.Descendants("Filter")
                .Where(f => string.Equals(Attr(f, "Category", string.Empty), "Category", StringComparison.OrdinalIgnoreCase))
                .Select(f => Attr(f, "Property", string.Empty))
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private List<string> CollectSuggestedParameterNamesFromXml()
        {
            if (_xml?.Root == null)
            {
                return new List<string>
                {
                    "ARG_SECTOR",
                    "ARG_NIVEL",
                    "ARG_SISTEMA",
                    "ARG_UNIFORMAT",
                    "ARG_CÓDIGO DE PARTIDA",
                    "ARG_UNIDAD DE PARTIDA",
                    "ARG_DESCRIPCIÓN DE PARTIDA",
                    "ARG_ESPECIALIDAD"
                };
            }

            List<string> names = _xml.Root.Descendants("Filter")
                .Where(f =>
                {
                    string fc = Attr(f, "Category", string.Empty);
                    return string.Equals(fc, "Parameter", StringComparison.OrdinalIgnoreCase) ||
                           string.Equals(fc, "APIParameter", StringComparison.OrdinalIgnoreCase) ||
                           string.Equals(fc, "HostParameter", StringComparison.OrdinalIgnoreCase);
                })
                .Select(f => Attr(f, "Property", string.Empty))
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (string fallback in new[]
            {
                "ARG_SECTOR",
                "ARG_NIVEL",
                "ARG_SISTEMA",
                "ARG_UNIFORMAT",
                "ARG_CÓDIGO DE PARTIDA",
                "ARG_UNIDAD DE PARTIDA",
                "ARG_DESCRIPCIÓN DE PARTIDA",
                "ARG_ESPECIALIDAD"
            })
            {
                if (!names.Any(v => string.Equals(v, fallback, StringComparison.OrdinalIgnoreCase)))
                {
                    names.Add(fallback);
                }
            }

            return names;
        }

        private void AttributesGridChanged(object? sender, EventArgs e)
        {
            if (_suppressGridEvents)
            {
                return;
            }

            ApplyAttributeGridToSelectedElement();
        }

        private void ApplyAttributeGridToSelectedElement()
        {
            XElement? element = GetSelectedElement();
            if (element == null)
            {
                return;
            }

            var attrs = new List<KeyValuePair<string, string>>();
            foreach (DataGridViewRow row in _attributesGrid.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                string key = (row.Cells[0].Value as string ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                string value = row.Cells[1].Value?.ToString() ?? string.Empty;
                attrs.Add(new KeyValuePair<string, string>(key, value));
            }

            element.RemoveAttributes();
            foreach (KeyValuePair<string, string> pair in attrs)
            {
                element.SetAttributeValue(pair.Key, pair.Value);
            }

            if (element.Name.LocalName == "MCSettings")
            {
                PopulateMetadataFields();
                UpdatePathUi();
            }

            UpdateTreeNodeText(_tree.SelectedNode);
            MarkDirty();
        }

        private void MetadataChanged(object? sender, EventArgs e)
        {
            if (_suppressMetadataEvents || _xml?.Root == null)
            {
                return;
            }

            XElement root = _xml.Root;
            root.SetAttributeValue("Name", _nameTextBox.Text.Trim());
            root.SetAttributeValue("Description", _descriptionTextBox.Text.Trim());
            root.SetAttributeValue("Author", _authorTextBox.Text.Trim());
            root.SetAttributeValue("Date", _dateTextBox.Text.Trim());
            root.SetAttributeValue("AllowRequired", _allowRequiredCheckBox.Checked ? "True" : "False");

            if (_tree.Nodes.Count > 0)
            {
                UpdateTreeNodeText(_tree.Nodes[0]);
            }

            UpdatePathUi();
            MarkDirty();
        }

        private void AddHeading()
        {
            if (_xml?.Root == null)
            {
                return;
            }

            XElement heading = CreateHeadingElement("Nuevo_Heading");
            _xml.Root.Add(heading);
            RepopulateAndSelect(heading);
            SetStatus("Heading agregado.");
        }

        private void AddSection()
        {
            XElement? selected = GetSelectedElement();
            XElement? heading = FindClosestAncestorByName(selected, "Heading");
            if (heading == null)
            {
                MessageBox.Show(this, "Selecciona un Heading para agregar una Section.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            XElement section = CreateSectionElement("Nueva_Section");
            heading.Add(section);
            RepopulateAndSelect(section);
            SetStatus("Section agregada.");
        }

        private void AddCheck()
        {
            XElement? selected = GetSelectedElement();
            XElement? section = FindClosestAncestorByName(selected, "Section");
            if (section == null)
            {
                MessageBox.Show(this, "Selecciona una Section para agregar un Check.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            XElement check = CreateCheckElement("Nuevo_Check");
            section.Add(check);
            RepopulateAndSelect(check);
            SetStatus("Check agregado.");
        }
        private void AddFilter()
        {
            XElement? selected = GetSelectedElement();
            XElement? check = FindClosestAncestorByName(selected, "Check");
            if (check == null)
            {
                MessageBox.Show(this, "Selecciona un Check para agregar un Filter.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            XElement filter = CreateFilterElement();
            check.Add(filter);
            RepopulateAndSelect(filter);
            SetStatus("Filter agregado.");
        }

        private void DuplicateSelectedNode()
        {
            XElement? selected = GetSelectedElement();
            if (selected == null || selected.Name.LocalName == "MCSettings")
            {
                return;
            }

            XElement? parent = selected.Parent;
            if (parent == null)
            {
                return;
            }

            XElement clone = new XElement(selected);
            ReassignIdsRecursively(clone);
            selected.AddAfterSelf(clone);

            RepopulateAndSelect(clone);
            SetStatus("Nodo duplicado.");
        }

        private void RenameSelectedNode()
        {
            XElement? selected = GetSelectedElement();
            if (selected == null)
            {
                return;
            }

            string elementName = selected.Name.LocalName;
            string attrName = elementName == "Heading"
                ? "HeadingText"
                : elementName == "Section"
                    ? "SectionName"
                    : elementName == "Check"
                        ? "CheckName"
                        : elementName == "Filter"
                            ? "Property"
                            : "Name";

            string current = Attr(selected, attrName, string.Empty);
            string? input = PromptForText(
                $"Renombrar {elementName}",
                $"Valor para atributo '{attrName}':",
                current);

            if (input == null)
            {
                return;
            }

            input = input.Trim();
            if (string.IsNullOrWhiteSpace(input))
            {
                MessageBox.Show(this, "El valor no puede ser vacío.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            selected.SetAttributeValue(attrName, input);
            UpdateTreeNodeText(_tree.SelectedNode);
            MarkDirty();
            SetStatus("Nodo renombrado.");
        }

        private void MoveSelectedNode(int direction)
        {
            if (direction == 0)
            {
                return;
            }

            XElement? selected = GetSelectedElement();
            if (selected == null || selected.Name.LocalName == "MCSettings")
            {
                return;
            }

            XElement? sibling = direction < 0
                ? selected.ElementsBeforeSelf().LastOrDefault(e => e.Name == selected.Name)
                : selected.ElementsAfterSelf().FirstOrDefault(e => e.Name == selected.Name);

            if (sibling == null)
            {
                return;
            }

            if (direction < 0)
            {
                selected.Remove();
                sibling.AddBeforeSelf(selected);
                SetStatus("Nodo movido hacia arriba.");
            }
            else
            {
                selected.Remove();
                sibling.AddAfterSelf(selected);
                SetStatus("Nodo movido hacia abajo.");
            }

            RepopulateAndSelect(selected);
        }

        private void GenerateChecksMatrix()
        {
            XElement? selected = GetSelectedElement();
            XElement? section = FindClosestAncestorByName(selected, "Section");
            if (section == null)
            {
                MessageBox.Show(this, "Selecciona una Section para generar checks.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            List<string> categories = CollectSuggestedCategoryProperties(section);
            List<string> parameters = CollectSuggestedParameterNames(section);

            if (categories.Count == 0)
            {
                categories = new List<string> { "OST_Walls", "OST_Doors", "OST_Windows", "OST_Floors", "OST_Rooms" };
            }

            if (parameters.Count == 0)
            {
                parameters = new List<string>
                {
                    "ARG_SECTOR",
                    "ARG_NIVEL",
                    "ARG_SISTEMA",
                    "ARG_UNIFORMAT",
                    "ARG_CÓDIGO DE PARTIDA",
                    "ARG_UNIDAD DE PARTIDA",
                    "ARG_DESCRIPCIÓN DE PARTIDA"
                };
            }

            if (!PromptMultiSelect("Generar matriz | Categorías", "Selecciona categorías (Property OST_*):", categories, out List<string> selectedCategories))
            {
                return;
            }

            if (!PromptMultiSelect("Generar matriz | Parámetros", "Selecciona parámetros a auditar:", parameters, out List<string> selectedParameters))
            {
                return;
            }

            if (selectedCategories.Count == 0 || selectedParameters.Count == 0)
            {
                MessageBox.Show(this, "Debes seleccionar al menos una categoría y un parámetro.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var existingNames = new HashSet<string>(
                section.Elements().Where(e => e.Name.LocalName == "Check")
                    .Select(c => Attr(c, "CheckName", string.Empty)),
                StringComparer.OrdinalIgnoreCase);

            int added = 0;
            int skipped = 0;
            XElement? firstAdded = null;

            foreach (string categoryProperty in selectedCategories)
            {
                string categoryToken = SanitizeToken(categoryProperty.Replace("OST_", string.Empty));
                foreach (string parameter in selectedParameters)
                {
                    string parameterToken = SanitizeToken(parameter);
                    string checkName = $"{categoryToken}_{parameterToken}_NoValue";
                    if (existingNames.Contains(checkName))
                    {
                        skipped++;
                        continue;
                    }

                    XElement check = CreateCheckElement(checkName);
                    check.SetAttributeValue("Description", $"PEB 4.11.3: parametro obligatorio '{parameter}' sin valor en categoria '{categoryToken}'");
                    check.SetAttributeValue("FailureMessage", $"Hay elementos en '{categoryToken}' con '{parameter}' sin valor");
                    check.RemoveNodes();

                    check.Add(new XElement("Filter",
                        new XAttribute("ID", Guid.NewGuid()),
                        new XAttribute("Operator", "And"),
                        new XAttribute("Category", "Category"),
                        new XAttribute("Property", categoryProperty),
                        new XAttribute("Condition", "Included"),
                        new XAttribute("Value", "True"),
                        new XAttribute("CaseInsensitive", "False"),
                        new XAttribute("Unit", "None"),
                        new XAttribute("UnitClass", "None"),
                        new XAttribute("FieldTitle", string.Empty),
                        new XAttribute("UserDefined", "False"),
                        new XAttribute("Validation", "None")));

                    check.Add(new XElement("Filter",
                        new XAttribute("ID", Guid.NewGuid()),
                        new XAttribute("Operator", "And"),
                        new XAttribute("Category", "Parameter"),
                        new XAttribute("Property", parameter),
                        new XAttribute("Condition", "HasNoValue"),
                        new XAttribute("Value", "True"),
                        new XAttribute("CaseInsensitive", "False"),
                        new XAttribute("Unit", "None"),
                        new XAttribute("UnitClass", "None"),
                        new XAttribute("FieldTitle", string.Empty),
                        new XAttribute("UserDefined", "False"),
                        new XAttribute("Validation", "None")));

                    check.Add(new XElement("Filter",
                        new XAttribute("ID", Guid.NewGuid()),
                        new XAttribute("Operator", "And"),
                        new XAttribute("Category", "TypeOrInstance"),
                        new XAttribute("Property", "Is Element Type"),
                        new XAttribute("Condition", "Equal"),
                        new XAttribute("Value", "False"),
                        new XAttribute("CaseInsensitive", "False"),
                        new XAttribute("Unit", "None"),
                        new XAttribute("UnitClass", "None"),
                        new XAttribute("FieldTitle", string.Empty),
                        new XAttribute("UserDefined", "False"),
                        new XAttribute("Validation", "None")));

                    section.Add(check);
                    existingNames.Add(checkName);
                    added++;
                    if (firstAdded == null)
                    {
                        firstAdded = check;
                    }
                }
            }

            if (added == 0)
            {
                MessageBox.Show(this, "No se agregaron checks nuevos. Todos ya existían.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Information);
                SetStatus("Matriz sin cambios (checks duplicados).");
                return;
            }

            RepopulateAndSelect(firstAdded ?? section);
            SetStatus($"Matriz generada: {added} check(s) nuevos, {skipped} omitidos.");
            MessageBox.Show(this, $"Matriz generada correctamente.{Environment.NewLine}Nuevos: {added}{Environment.NewLine}Omitidos: {skipped}", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void DeleteSelectedNode()
        {
            XElement? selected = GetSelectedElement();
            if (selected == null || selected.Name.LocalName == "MCSettings")
            {
                return;
            }

            DialogResult confirm = MessageBox.Show(
                this,
                "Se eliminara el nodo seleccionado y su contenido. Deseas continuar?",
                "JeiAudit",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes)
            {
                return;
            }

            XElement? parent = selected.Parent;
            selected.Remove();
            RepopulateAndSelect(parent);
            SetStatus("Nodo eliminado.");
        }

        private void RepopulateAndSelect(XElement? toSelect)
        {
            PopulateTree();
            if (toSelect != null)
            {
                SelectNodeByElementReference(toSelect);
            }
            MarkDirty();
        }

        private void SelectNodeByElementReference(XElement element)
        {
            foreach (TreeNode node in _tree.Nodes)
            {
                TreeNode? found = FindNodeByElement(node, element);
                if (found != null)
                {
                    _tree.SelectedNode = found;
                    found.EnsureVisible();
                    return;
                }
            }
        }

        private static TreeNode? FindNodeByElement(TreeNode node, XElement element)
        {
            if (ReferenceEquals(node.Tag, element))
            {
                return node;
            }

            foreach (TreeNode child in node.Nodes)
            {
                TreeNode? found = FindNodeByElement(child, element);
                if (found != null)
                {
                    return found;
                }
            }

            return null;
        }

        private bool SaveXml(bool forceSaveAs, bool copyOnly)
        {
            if (_xml?.Root == null)
            {
                MessageBox.Show(this, "No hay checkset cargado.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            string targetPath = SelectedXmlPath;
            if (forceSaveAs || string.IsNullOrWhiteSpace(targetPath))
            {
                string suggested = string.IsNullOrWhiteSpace(SelectedXmlPath)
                    ? "MCSettings_Nuevo_R2024.xml"
                    : Path.GetFileNameWithoutExtension(SelectedXmlPath) + (copyOnly ? "_copy" : string.Empty) + ".xml";

                using (var dialog = new SaveFileDialog())
                {
                    dialog.Title = copyOnly ? "Copiar checkset como..." : "Guardar checkset como...";
                    dialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
                    dialog.FileName = suggested;
                    dialog.OverwritePrompt = true;
                    dialog.AddExtension = true;
                    dialog.DefaultExt = "xml";

                    if (!string.IsNullOrWhiteSpace(SelectedXmlPath) && File.Exists(SelectedXmlPath))
                    {
                        dialog.InitialDirectory = Path.GetDirectoryName(SelectedXmlPath);
                    }

                    if (dialog.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(dialog.FileName))
                    {
                        return false;
                    }

                    targetPath = dialog.FileName;
                }
            }

            try
            {
                _xml.Save(targetPath, SaveOptions.None);
                if (copyOnly)
                {
                    MessageBox.Show(this, "Copia de checkset guardada.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    SetStatus("Copia de checkset guardada.");
                }
                else
                {
                    SelectedXmlPath = targetPath;
                    PluginState.SaveLastXmlPath(SelectedXmlPath);
                    _hasChanges = false;
                    UpdatePathUi();
                    MessageBox.Show(this, "Checkset guardado correctamente.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    SetStatus("Checkset guardado.");
                }

                RefreshUiState();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    $"No se pudo guardar el XML.{Environment.NewLine}{ex.Message}",
                    "JeiAudit",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
        }

        private void OnFormClosing(object? sender, FormClosingEventArgs e)
        {
            if (!_hasChanges)
            {
                DialogResult = DialogResult.OK;
                return;
            }

            DialogResult confirm = MessageBox.Show(
                this,
                "Hay cambios sin guardar. Deseas guardar antes de cerrar?",
                "JeiAudit",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (confirm == DialogResult.Cancel)
            {
                e.Cancel = true;
                return;
            }

            if (confirm == DialogResult.Yes && !SaveXml(forceSaveAs: false, copyOnly: false))
            {
                e.Cancel = true;
                return;
            }

            DialogResult = DialogResult.OK;
        }

        private bool ConfirmCloseOrDiscardChanges()
        {
            if (!_hasChanges)
            {
                return true;
            }

            DialogResult confirm = MessageBox.Show(
                this,
                "Hay cambios sin guardar. Deseas continuar sin guardarlos?",
                "JeiAudit",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            return confirm == DialogResult.Yes;
        }

        private void RefreshUiState()
        {
            bool hasXml = _xml?.Root != null;
            XElement? selected = GetSelectedElement();
            string selectedKind = selected?.Name.LocalName ?? string.Empty;
            bool canReorder = selected != null && selectedKind != "MCSettings";

            _saveButton.Enabled = hasXml;
            _saveAsButton.Enabled = hasXml;
            _copyButton.Enabled = hasXml;
            _addHeadingButton.Enabled = hasXml;
            _deleteButton.Enabled = canReorder;
            _addSectionButton.Enabled = FindClosestAncestorByName(selected, "Heading") != null;
            _addCheckButton.Enabled = FindClosestAncestorByName(selected, "Section") != null;
            _addFilterButton.Enabled = FindClosestAncestorByName(selected, "Check") != null;
            _duplicateButton.Enabled = canReorder;
            _renameButton.Enabled = canReorder;
            _moveUpButton.Enabled = canReorder;
            _moveDownButton.Enabled = canReorder;
            _matrixButton.Enabled = FindClosestAncestorByName(selected, "Section") != null;
            _findNextButton.Enabled = hasXml;
            _searchTextBox.Enabled = hasXml;

            _nameTextBox.Enabled = hasXml;
            _descriptionTextBox.Enabled = hasXml;
            _authorTextBox.Enabled = hasXml;
            _dateTextBox.Enabled = hasXml;
            _allowRequiredCheckBox.Enabled = hasXml;
            _attributesGrid.Enabled = selected != null;
        }

        private void MarkDirty()
        {
            _hasChanges = true;
            RefreshUiState();
        }

        private void UpdatePathUi()
        {
            if (string.IsNullOrWhiteSpace(SelectedXmlPath))
            {
                _statusLabel.Text = "Nuevo";
                _statusLabel.ForeColor = Color.FromArgb(72, 72, 72);
                _pathLabel.Text = "(Sin archivo guardado)";
            }
            else
            {
                _statusLabel.Text = "Abierto";
                _statusLabel.ForeColor = Color.FromArgb(72, 72, 72);
                _pathLabel.Text = SelectedXmlPath;
            }
        }

        private XElement? GetSelectedElement()
        {
            return _tree.SelectedNode?.Tag as XElement;
        }

        private static XElement? FindClosestAncestorByName(XElement? element, string targetName)
        {
            XElement? current = element;
            while (current != null)
            {
                if (string.Equals(current.Name.LocalName, targetName, StringComparison.OrdinalIgnoreCase))
                {
                    return current;
                }

                current = current.Parent;
            }

            return null;
        }

        private static string BuildElementPath(XElement element)
        {
            var stack = new Stack<string>();
            XElement? current = element;
            while (current != null)
            {
                stack.Push(BuildNodeText(current));
                current = current.Parent;
            }

            return string.Join(" > ", stack);
        }

        private void UpdateTreeNodeText(TreeNode? node)
        {
            if (node == null || !(node.Tag is XElement element))
            {
                return;
            }

            node.Text = BuildNodeText(element);
        }

        private static void ReassignIdsRecursively(XElement element)
        {
            if (element.Attribute("ID") != null)
            {
                element.SetAttributeValue("ID", Guid.NewGuid().ToString());
            }

            foreach (XElement child in element.Elements())
            {
                ReassignIdsRecursively(child);
            }
        }

        private void FindNextInTree()
        {
            string term = (_searchTextBox?.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(term))
            {
                SetStatus("Escribe texto para buscar.");
                return;
            }

            List<TreeNode> nodes = FlattenTreeNodes().ToList();
            if (nodes.Count == 0)
            {
                return;
            }

            int start = _searchCursor + 1;
            if (start >= nodes.Count)
            {
                start = 0;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                int idx = (start + i) % nodes.Count;
                TreeNode node = nodes[idx];
                if ((node.Text ?? string.Empty).IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    _searchCursor = idx;
                    _tree.SelectedNode = node;
                    node.EnsureVisible();
                    SetStatus($"Encontrado: {node.Text}");
                    return;
                }
            }

            SetStatus("Sin resultados para búsqueda.");
        }

        private IEnumerable<TreeNode> FlattenTreeNodes()
        {
            foreach (TreeNode node in _tree.Nodes)
            {
                foreach (TreeNode item in Flatten(node))
                {
                    yield return item;
                }
            }
        }

        private static IEnumerable<TreeNode> Flatten(TreeNode node)
        {
            yield return node;
            foreach (TreeNode child in node.Nodes)
            {
                foreach (TreeNode item in Flatten(child))
                {
                    yield return item;
                }
            }
        }

        private static string? PromptForText(string title, string label, string initialValue)
        {
            using (var form = new Form())
            {
                form.Text = title;
                form.StartPosition = FormStartPosition.CenterParent;
                form.Width = 620;
                form.Height = 180;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MinimizeBox = false;
                form.MaximizeBox = false;
                form.ShowInTaskbar = false;

                var lbl = new Label
                {
                    Parent = form,
                    Left = 14,
                    Top = 14,
                    Width = 570,
                    Height = 24,
                    Text = label,
                    Font = new Font("Segoe UI", 10f, FontStyle.Regular)
                };

                var txt = new TextBox
                {
                    Parent = form,
                    Left = 14,
                    Top = 44,
                    Width = 572,
                    Height = 26,
                    Text = initialValue ?? string.Empty,
                    Font = new Font("Segoe UI", 10f, FontStyle.Regular)
                };

                var ok = new Button
                {
                    Parent = form,
                    Text = "Aceptar",
                    Width = 100,
                    Height = 32,
                    Left = 378,
                    Top = 86,
                    DialogResult = DialogResult.OK
                };

                var cancel = new Button
                {
                    Parent = form,
                    Text = "Cancelar",
                    Width = 100,
                    Height = 32,
                    Left = 486,
                    Top = 86,
                    DialogResult = DialogResult.Cancel
                };

                form.AcceptButton = ok;
                form.CancelButton = cancel;

                DialogResult result = form.ShowDialog();
                return result == DialogResult.OK ? txt.Text : null;
            }
        }

        private bool PromptMultiSelect(
            string title,
            string caption,
            List<string> sourceItems,
            out List<string> selectedItems)
        {
            selectedItems = new List<string>();
            List<string> items = sourceItems
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (items.Count == 0)
            {
                return false;
            }

            using (var form = new Form())
            {
                form.Text = title;
                form.StartPosition = FormStartPosition.CenterParent;
                form.Width = 540;
                form.Height = 640;
                form.MinimumSize = new Size(500, 520);
                form.BackColor = Color.White;

                form.Controls.Add(new Label
                {
                    Text = caption,
                    Left = 12,
                    Top = 12,
                    Width = 500,
                    Height = 24,
                    Font = new Font("Segoe UI", 10f, FontStyle.Bold)
                });

                var list = new CheckedListBox
                {
                    Parent = form,
                    Left = 12,
                    Top = 40,
                    Width = 500,
                    Height = 500,
                    CheckOnClick = true,
                    Font = new Font("Segoe UI", 9.5f, FontStyle.Regular)
                };

                foreach (string item in items)
                {
                    list.Items.Add(item, true);
                }

                var selectAll = new Button
                {
                    Parent = form,
                    Text = "Todos",
                    Left = 12,
                    Top = 548,
                    Width = 90,
                    Height = 30
                };
                selectAll.Click += (_, _) =>
                {
                    for (int i = 0; i < list.Items.Count; i++)
                    {
                        list.SetItemChecked(i, true);
                    }
                };

                var clearAll = new Button
                {
                    Parent = form,
                    Text = "Ninguno",
                    Left = 108,
                    Top = 548,
                    Width = 90,
                    Height = 30
                };
                clearAll.Click += (_, _) =>
                {
                    for (int i = 0; i < list.Items.Count; i++)
                    {
                        list.SetItemChecked(i, false);
                    }
                };

                var ok = new Button
                {
                    Parent = form,
                    Text = "Aceptar",
                    Left = 322,
                    Top = 548,
                    Width = 90,
                    Height = 30,
                    DialogResult = DialogResult.OK
                };

                var cancel = new Button
                {
                    Parent = form,
                    Text = "Cancelar",
                    Left = 422,
                    Top = 548,
                    Width = 90,
                    Height = 30,
                    DialogResult = DialogResult.Cancel
                };

                form.AcceptButton = ok;
                form.CancelButton = cancel;

                if (form.ShowDialog(this) != DialogResult.OK)
                {
                    return false;
                }

                selectedItems = list.CheckedItems.Cast<object>()
                    .Select(v => v?.ToString() ?? string.Empty)
                    .Where(v => !string.IsNullOrWhiteSpace(v))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                return true;
            }
        }

        private List<string> CollectSuggestedCategoryProperties(XElement section)
        {
            return section.Elements()
                .Where(e => e.Name.LocalName == "Check")
                .SelectMany(c => c.Elements().Where(f => f.Name.LocalName == "Filter"))
                .Where(f => string.Equals(Attr(f, "Category", string.Empty), "Category", StringComparison.OrdinalIgnoreCase))
                .Select(f => Attr(f, "Property", string.Empty))
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private List<string> CollectSuggestedParameterNames(XElement section)
        {
            return section.Elements()
                .Where(e => e.Name.LocalName == "Check")
                .SelectMany(c => c.Elements().Where(f => f.Name.LocalName == "Filter"))
                .Where(f =>
                {
                    string fc = Attr(f, "Category", string.Empty);
                    return string.Equals(fc, "Parameter", StringComparison.OrdinalIgnoreCase) ||
                           string.Equals(fc, "APIParameter", StringComparison.OrdinalIgnoreCase);
                })
                .Select(f => Attr(f, "Property", string.Empty))
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string SanitizeToken(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "ITEM";
            }

            var chars = value.Trim()
                .Select(c => char.IsLetterOrDigit(c) ? c : '_')
                .ToArray();
            string normalized = new string(chars);
            while (normalized.Contains("__", StringComparison.Ordinal))
            {
                normalized = normalized.Replace("__", "_");
            }
            return normalized.Trim('_');
        }

        private void SetStatus(string text)
        {
            if (_footerStatusLabel != null)
            {
                _footerStatusLabel.Text = string.IsNullOrWhiteSpace(text) ? "Listo." : text;
            }
        }

        private static XElement CreateHeadingElement(string headingText)
        {
            return new XElement("Heading",
                new XAttribute("ID", Guid.NewGuid()),
                new XAttribute("HeadingText", headingText),
                new XAttribute("Description", string.Empty),
                new XAttribute("IsChecked", "True"));
        }

        private static XElement CreateSectionElement(string sectionName)
        {
            return new XElement("Section",
                new XAttribute("ID", Guid.NewGuid()),
                new XAttribute("SectionName", sectionName),
                new XAttribute("Title", string.Empty),
                new XAttribute("IsChecked", "True"),
                new XAttribute("Description", string.Empty));
        }

        private static XElement CreateCheckElement(string checkName)
        {
            return new XElement("Check",
                new XAttribute("ID", Guid.NewGuid()),
                new XAttribute("CheckName", checkName),
                new XAttribute("Description", "Nuevo check creado con editor JeiAudit"),
                new XAttribute("FailureMessage", "El check ha fallado."),
                new XAttribute("ResultCondition", "FailMatchingElements"),
                new XAttribute("CheckType", "Custom"),
                new XAttribute("IsRequired", "True"),
                new XAttribute("IsChecked", "True"));
        }

        private static XElement CreateFilterElement()
        {
            return new XElement("Filter",
                new XAttribute("ID", Guid.NewGuid()),
                new XAttribute("Operator", "And"),
                new XAttribute("Category", "Category"),
                new XAttribute("Property", "OST_Walls"),
                new XAttribute("Condition", "Included"),
                new XAttribute("Value", "True"),
                new XAttribute("CaseInsensitive", "False"),
                new XAttribute("Unit", "None"),
                new XAttribute("UnitClass", "None"),
                new XAttribute("FieldTitle", string.Empty),
                new XAttribute("UserDefined", "False"),
                new XAttribute("Validation", "None"));
        }

        private static string Attr(XElement? element, string name, string fallback)
        {
            if (element == null)
            {
                return fallback;
            }

            XAttribute? attribute = element.Attributes()
                .FirstOrDefault(a => string.Equals(a.Name.LocalName, name, StringComparison.Ordinal));
            string value = (attribute?.Value ?? string.Empty).Trim();
            return value.Length == 0 ? fallback : value;
        }

        private static bool IsTrue(string value)
        {
            return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) || value == "1";
        }
    }
}
