using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace JeiAudit
{
    internal static class UiTheme
    {
        internal static readonly Color WindowBackground = Color.FromArgb(245, 247, 250);
        internal static readonly Color Surface = Color.White;
        internal static readonly Color SurfaceMuted = Color.FromArgb(248, 250, 252);
        internal static readonly Color Border = Color.FromArgb(219, 225, 232);

        internal static readonly Color HeaderBackground = Color.FromArgb(27, 41, 63);
        internal static readonly Color HeaderTitle = Color.White;
        internal static readonly Color HeaderSubtitle = Color.FromArgb(198, 208, 223);

        internal static readonly Color TextPrimary = Color.FromArgb(30, 41, 59);
        internal static readonly Color TextSecondary = Color.FromArgb(88, 102, 120);

        internal static readonly Color Accent = Color.FromArgb(32, 114, 233);
        internal static readonly Color AccentHover = Color.FromArgb(26, 99, 207);
        internal static readonly Color AccentPressed = Color.FromArgb(22, 86, 182);
        internal static readonly Color AccentSoft = Color.FromArgb(232, 241, 253);

        internal static readonly Color NeutralButton = Color.FromArgb(245, 248, 252);
        internal static readonly Color NeutralButtonHover = Color.FromArgb(237, 242, 248);
        internal static readonly Color NeutralButtonPressed = Color.FromArgb(225, 234, 245);

        internal static void EnableSmoothRendering(Form form)
        {
            if (form == null)
            {
                return;
            }

            try
            {
                typeof(Control)
                    .GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)
                    ?.SetValue(form, true, null);

                typeof(Control)
                    .GetProperty("ResizeRedraw", BindingFlags.Instance | BindingFlags.NonPublic)
                    ?.SetValue(form, true, null);
            }
            catch
            {
                // Ignore if the runtime blocks protected reflection access.
            }
        }

        internal static void StyleCard(Panel panel)
        {
            panel.BackColor = Surface;
            panel.BorderStyle = BorderStyle.None;
            panel.Paint += (_, e) =>
            {
                using (var pen = new Pen(Border))
                {
                    Rectangle rect = new Rectangle(0, 0, panel.Width - 1, panel.Height - 1);
                    e.Graphics.DrawRectangle(pen, rect);
                }
            };
        }

        internal static void StyleToolbarButton(Button button, bool primary = false)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 1;
            button.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            button.Cursor = Cursors.Hand;
            button.UseVisualStyleBackColor = false;

            Color normalBack = primary ? Accent : NeutralButton;
            Color hoverBack = primary ? AccentHover : NeutralButtonHover;
            Color pressedBack = primary ? AccentPressed : NeutralButtonPressed;
            Color borderColor = primary ? Accent : Border;
            Color foreColor = primary ? Color.White : TextPrimary;

            button.BackColor = normalBack;
            button.ForeColor = foreColor;
            button.FlatAppearance.BorderColor = borderColor;

            button.MouseEnter += (_, _) => button.BackColor = hoverBack;
            button.MouseLeave += (_, _) => button.BackColor = normalBack;
            button.MouseDown += (_, _) => button.BackColor = pressedBack;
            button.MouseUp += (_, _) => button.BackColor = button.ClientRectangle.Contains(button.PointToClient(Cursor.Position)) ? hoverBack : normalBack;
        }

        internal static void StyleFooterButton(Button button, bool primary = false)
        {
            StyleToolbarButton(button, primary);
            button.Font = new Font("Segoe UI", 10.5f, FontStyle.Bold);
        }
    }
}
