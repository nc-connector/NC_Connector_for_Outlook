/**
 * Copyright (c) 2026 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class FooterButtonLayoutHelper
    {
        internal const int DefaultHorizontalPadding = 12;
        internal const int DefaultBottomPadding = 12;
        internal const int DefaultSpacing = 12;

        internal static void ApplyButtonSize(Button button, out int minWidth)
        {
            minWidth = 96;
            // Layout helper is used in dynamic form states; null means "skip sizing" rather than fail.
            if (button == null)
            {
                return;
            }

            Size textSize = TextRenderer.MeasureText(button.Text ?? string.Empty, button.Font);
            int width = Math.Max(132, textSize.Width + 50);
            // Keep minimum width close to preferred width to avoid text wrap/distortion while resizing.
            minWidth = Math.Max(124, width - 6);
            int height = Math.Max(32, textSize.Height + 12);

            button.Size = new Size(width, height);
            button.TextAlign = ContentAlignment.MiddleCenter;
            button.AutoSize = false;
        }

        internal static int LayoutCentered(Control container, IList<Button> buttons, int horizontalPadding, int bottomPadding, int spacing)
        {
            return LayoutCentered(container, buttons, horizontalPadding, bottomPadding, spacing, false);
        }

        internal static int LayoutCentered(Control container, IList<Button> buttons, int horizontalPadding, int bottomPadding, int spacing, bool uniformWidths)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (container == null || buttons == null || buttons.Count == 0)
            {
                return 0;
            }

            var minWidths = new Dictionary<Button, int>();
            int buttonHeight = 0;
            foreach (Button button in buttons)
            {
                int minWidth;
                ApplyButtonSize(button, out minWidth);
                minWidths[button] = minWidth;
                buttonHeight = Math.Max(buttonHeight, button.Height);
            }

            if (uniformWidths)
            {
                int commonWidth = 0;
                int commonMinWidth = 0;
                foreach (Button button in buttons)
                {
                    if (button.Width > commonWidth)
                    {
                        commonWidth = button.Width;
                    }

                    int buttonMinWidth;
                    if (minWidths.TryGetValue(button, out buttonMinWidth) && buttonMinWidth > commonMinWidth)
                    {
                        commonMinWidth = buttonMinWidth;
                    }
                }

                foreach (Button button in buttons)
                {
                    button.Width = commonWidth;
                    minWidths[button] = commonMinWidth;
                }
            }

            int availableWidth = Math.Max(0, container.ClientSize.Width - (horizontalPadding * 2));
            if (availableWidth <= 0)
            {
                return 0;
            }

            int safeSpacing = Math.Max(0, spacing);
            int totalWidth = CalculateTotalWidth(buttons, safeSpacing);
            while (safeSpacing > 0 && totalWidth > availableWidth)
            {
                safeSpacing--;
                totalWidth = CalculateTotalWidth(buttons, safeSpacing);
            }

            if (totalWidth > availableWidth)
            {
                int overflow = totalWidth - availableWidth;
                bool changed = true;
                while (overflow > 0 && changed)
                {
                    changed = false;
                    for (int i = 0; i < buttons.Count && overflow > 0; i++)
                    {
                        Button button = buttons[i];
                        int minWidth = minWidths.ContainsKey(button) ? minWidths[button] : 1;
                        if (button.Width > minWidth)
                        {
                            button.Width -= 1;
                            overflow -= 1;
                            changed = true;
                        }
                    }
                }

                totalWidth = CalculateTotalWidth(buttons, safeSpacing);
            }

            int top = Math.Max(horizontalPadding, container.ClientSize.Height - buttonHeight - bottomPadding);
            int left = horizontalPadding + Math.Max(0, (availableWidth - totalWidth) / 2);
            for (int i = 0; i < buttons.Count; i++)
            {
                Button button = buttons[i];
                button.Location = new Point(left, top);
                left += button.Width + safeSpacing;
            }

            return totalWidth + (horizontalPadding * 2);
        }

        private static int CalculateTotalWidth(IList<Button> buttons, int spacing)
        {
            int total = 0;
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (buttons == null)
            {
                return total;
            }

            for (int i = 0; i < buttons.Count; i++)
            {
                total += buttons[i].Width;
            }

            if (buttons.Count > 1)
            {
                total += spacing * (buttons.Count - 1);
            }

            return total;
        }
    }
}
