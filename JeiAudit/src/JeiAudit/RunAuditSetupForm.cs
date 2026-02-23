using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace JeiAudit
{
    internal sealed class RunAuditChecksetMetadata
    {
        public string Title { get; set; } = "-";
        public string Date { get; set; } = "-";
        public string Author { get; set; } = "-";
        public string Description { get; set; } = "-";
    }

    internal sealed class RunAuditModelItem
    {
        public Autodesk.Revit.DB.Document Document { get; set; } = null!;
        public string DisplayName { get; set; } = string.Empty;
        public string Path { get; set; } = string.Empty;
        public bool IsLink { get; set; }
        public bool IsSelected { get; set; }
    }

    internal sealed class RunAuditSetupForm : Form
    {
        private readonly List<RunAuditModelItem> _models;
        private ListView _modelsListView = null!;

        internal RunAuditSetupForm(RunAuditChecksetMetadata metadata, List<RunAuditModelItem> models)
        {
            _models = models;
            Text = "Herramienta de auditor\u00EDa JeiAudit | Ejecutar | Desarrollado por Jason Rojas Estrada - Coordinador BIM";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 840;
            Height = 700;
            MinimumSize = new Size(760, 620);
            BackColor = UiTheme.WindowBackground;
            UiTheme.EnableSmoothRendering(this);

            Panel header = BuildHeader();
            Controls.Add(header);

            Panel footer = BuildFooter();
            Controls.Add(footer);

            Panel content = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(22, 16, 22, 22),
                BackColor = UiTheme.WindowBackground
            };
            Controls.Add(content);

            Panel runCard = BuildCard(content);
            runCard.Dock = DockStyle.Fill;

            BuildRunBody(runCard);

            Panel metadataCard = BuildCard(content);
            metadataCard.Dock = DockStyle.Top;
            metadataCard.Height = 190;
            BuildMetadata(metadataCard, metadata);
        }

        internal List<RunAuditModelItem> GetSelectedModels()
        {
            return _models.Where(v => v.IsSelected).ToList();
        }

        private Panel BuildHeader()
        {
            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 72,
                BackColor = UiTheme.HeaderBackground
            };
            header.Controls.Add(new Label
            {
                Text = "Herramienta de auditor\u00EDa JeiAudit",
                ForeColor = UiTheme.HeaderTitle,
                Font = new Font("Segoe UI", 15f, FontStyle.Bold),
                AutoSize = true,
                Left = 54,
                Top = 4
            });
            header.Controls.Add(new Label
            {
                Text = "Desarrollado por Jason Rojas Estrada - Coordinador BIM, Inspirada en herramientas de Autodesk",
                ForeColor = UiTheme.HeaderSubtitle,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular),
                AutoSize = true,
                Left = 54,
                Top = 34
            });
            var logo = new Panel
            {
                Parent = header,
                Width = 22,
                Height = 22,
                Left = 16,
                Top = 11,
                BackColor = UiTheme.AccentSoft
            };
            logo.Paint += (_, e) =>
            {
                using (var pen = new Pen(UiTheme.Accent, 2))
                {
                    e.Graphics.DrawRectangle(pen, 1, 1, 19, 19);
                }
            };
            return header;
        }

        private Panel BuildFooter()
        {
            var footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 58,
                BackColor = UiTheme.Surface
            };
            footer.Paint += (_, e) =>
            {
                using (var pen = new Pen(UiTheme.Border))
                {
                    e.Graphics.DrawLine(pen, 0, 0, footer.Width, 0);
                }
            };

            var table = new TableLayoutPanel
            {
                Parent = footer,
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                Padding = new Padding(12, 10, 12, 10)
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));

            Button cancel = BuildFooterButton("Cancelar", primary: false);
            cancel.Click += (_, _) =>
            {
                DialogResult = DialogResult.Cancel;
                Close();
            };
            table.Controls.Add(cancel, 0, 0);

            Button run = BuildFooterButton("Ejecutar comprobacion", primary: true);
            run.Click += (_, _) => ExecuteRun();
            table.Controls.Add(run, 1, 0);

            return footer;
        }

        private static Button BuildFooterButton(string text, bool primary)
        {
            var button = new Button
            {
                Dock = DockStyle.Fill,
                Text = text,
                Margin = new Padding(6, 0, 6, 0)
            };
            UiTheme.StyleFooterButton(button, primary);
            return button;
        }

        private static Panel BuildCard(Control parent)
        {
            var panel = new Panel
            {
                Parent = parent,
                BackColor = UiTheme.Surface,
                Margin = new Padding(0, 10, 0, 0)
            };
            UiTheme.StyleCard(panel);
            return panel;
        }

        private static void BuildMetadata(Panel card, RunAuditChecksetMetadata metadata)
        {
            var icon = new Panel
            {
                Parent = card,
                Width = 128,
                Dock = DockStyle.Left,
                BackColor = UiTheme.Surface
            };
            icon.Paint += (_, e) =>
            {
                var top = new[] { new Point(18, 46), new Point(46, 18), new Point(112, 18), new Point(84, 46) };
                var left = new[] { new Point(18, 46), new Point(84, 46), new Point(84, 112), new Point(18, 112) };
                var right = new[] { new Point(84, 46), new Point(112, 18), new Point(112, 84), new Point(84, 112) };

                using (var topBrush = new SolidBrush(Color.FromArgb(231, 231, 231)))
                using (var leftBrush = new SolidBrush(Color.FromArgb(226, 202, 122)))
                using (var rightBrush = new SolidBrush(Color.FromArgb(197, 200, 205)))
                using (var pen = new Pen(Color.FromArgb(107, 107, 107), 3))
                {
                    e.Graphics.FillPolygon(topBrush, top);
                    e.Graphics.FillPolygon(leftBrush, left);
                    e.Graphics.FillPolygon(rightBrush, right);
                    e.Graphics.DrawPolygon(pen, top);
                    e.Graphics.DrawPolygon(pen, left);
                    e.Graphics.DrawPolygon(pen, right);
                }
            };

            var text = new Panel
            {
                Parent = card,
                Dock = DockStyle.Fill,
                Padding = new Padding(12, 12, 12, 8),
                BackColor = UiTheme.Surface
            };
            BuildMetadataRow(text, "Titulo", metadata.Title, 12);
            BuildMetadataRow(text, "Fecha", metadata.Date, 50);
            BuildMetadataRow(text, "Autor", metadata.Author, 88);
            BuildMetadataRow(text, "Descripcion", metadata.Description, 126);
        }

        private static void BuildMetadataRow(Control parent, string label, string value, int top)
        {
            parent.Controls.Add(new Label
            {
                Text = label,
                Font = new Font("Segoe UI", 12f, FontStyle.Regular),
                AutoSize = true,
                Left = 8,
                Top = top
            });
            parent.Controls.Add(new Label
            {
                Text = value,
                Font = new Font("Segoe UI", 10.5f, FontStyle.Regular),
                AutoSize = false,
                Left = 110,
                Top = top + 5,
                Width = 580,
                Height = 28,
                ForeColor = UiTheme.TextSecondary,
                AutoEllipsis = true
            });
        }

        private void BuildRunBody(Panel runCard)
        {
            var body = new SplitContainer
            {
                Parent = runCard,
                Dock = DockStyle.Fill,
                SplitterDistance = 276,
                BackColor = UiTheme.WindowBackground
            };
            body.Panel1.BackColor = UiTheme.Surface;
            body.Panel2.BackColor = UiTheme.Surface;

            var actionsPanel = new Panel
            {
                Parent = body.Panel1,
                Dock = DockStyle.Fill,
                Padding = new Padding(14, 14, 14, 14),
                BackColor = UiTheme.Surface
            };

            Button addBtn = BuildActionButton("Agregar modelos", 0);
            addBtn.Click += (_, _) => SetAllModelsChecked(true);
            actionsPanel.Controls.Add(addBtn);

            Button clearBtn = BuildActionButton("Quitar todos los modelos", 1);
            clearBtn.Click += (_, _) => SetAllModelsChecked(false);
            actionsPanel.Controls.Add(clearBtn);

            Button checkLinksBtn = BuildActionButton("Comprobar todos los enlaces", 2);
            checkLinksBtn.Click += (_, _) => SetLinksChecked(true);
            actionsPanel.Controls.Add(checkLinksBtn);

            Button uncheckLinksBtn = BuildActionButton("Desmarque todos los enlaces", 3);
            uncheckLinksBtn.Click += (_, _) => SetLinksChecked(false);
            actionsPanel.Controls.Add(uncheckLinksBtn);

            var rightPanel = new Panel
            {
                Parent = body.Panel2,
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 14, 14, 14),
                BackColor = UiTheme.Surface
            };
            rightPanel.Controls.Add(new Label
            {
                Text = "Modelos a comprobar",
                Font = new Font("Segoe UI", 13f, FontStyle.Bold),
                AutoSize = true,
                Left = 2,
                Top = 2,
                ForeColor = UiTheme.TextPrimary
            });

            _modelsListView = new ListView
            {
                Parent = rightPanel,
                CheckBoxes = true,
                View = View.Details,
                FullRowSelect = true,
                GridLines = false,
                Left = 0,
                Top = 40,
                Width = rightPanel.ClientSize.Width,
                Height = rightPanel.ClientSize.Height - 44,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 10f, FontStyle.Regular),
                BorderStyle = BorderStyle.FixedSingle
            };
            _modelsListView.Columns.Add("Modelo", 320);
            _modelsListView.Columns.Add("Ruta", 520);
            _modelsListView.ItemChecked += ModelItemChecked;

            PopulateModels();
        }

        private Button BuildActionButton(string text, int index)
        {
            var button = new Button
            {
                Text = text,
                Width = 238,
                Height = 54,
                Left = 0,
                Top = index * 66,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold)
            };
            UiTheme.StyleToolbarButton(button, primary: false);
            return button;
        }

        private void PopulateModels()
        {
            _modelsListView.Items.Clear();
            foreach (RunAuditModelItem model in _models)
            {
                string name = model.IsLink ? "[Link] " + model.DisplayName : model.DisplayName;
                string path = string.IsNullOrWhiteSpace(model.Path) ? "(sin ruta)" : model.Path;
                var item = new ListViewItem(name) { Tag = model, Checked = model.IsSelected };
                item.SubItems.Add(path);
                _modelsListView.Items.Add(item);
            }
        }

        private void ModelItemChecked(object? sender, ItemCheckedEventArgs e)
        {
            if (e.Item.Tag is RunAuditModelItem model)
            {
                model.IsSelected = e.Item.Checked;
            }
        }

        private void SetAllModelsChecked(bool value)
        {
            foreach (ListViewItem item in _modelsListView.Items)
            {
                item.Checked = value;
            }
        }

        private void SetLinksChecked(bool value)
        {
            foreach (ListViewItem item in _modelsListView.Items)
            {
                if (item.Tag is RunAuditModelItem model && model.IsLink)
                {
                    item.Checked = value;
                }
            }
        }

        private void ExecuteRun()
        {
            if (_models.All(v => !v.IsSelected))
            {
                MessageBox.Show(this, "Selecciona al menos un modelo para ejecutar la comprobacion.", "JeiAudit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
