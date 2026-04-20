/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Mirrors lock/explanation tooltips from unavailable controls onto a small active
     * hint glyph so policy/admin/backend hints remain reachable in WinForms.
     */
    internal sealed class DisabledControlTooltipHintHelper
    {
        private const int HintSize = 16;
        private const int HintSpacing = 6;

        private readonly ToolTip _toolTip;
        private readonly Dictionary<Control, Label> _hintLabels = new Dictionary<Control, Label>();
        private readonly Dictionary<Control, Control> _anchorOverridesByPrimary = new Dictionary<Control, Control>();
        private readonly Dictionary<Control, Control[]> _fallbackTargetsByPrimary = new Dictionary<Control, Control[]>();
        private readonly Dictionary<Control, bool> _showHintByPrimary = new Dictionary<Control, bool>();
        private readonly HashSet<Control> _trackedControls = new HashSet<Control>();

        internal DisabledControlTooltipHintHelper(ToolTip toolTip)
        {
            if (toolTip == null)
            {
                throw new ArgumentNullException("toolTip");
            }

            _toolTip = toolTip;
        }

        internal void Apply(Control primary, string text, params Control[] fallbackTargets)
        {
            Apply(primary, text, false, null, fallbackTargets);
        }

        internal void Apply(Control primary, string text, bool showHint, params Control[] fallbackTargets)
        {
            Apply(primary, text, showHint, null, fallbackTargets);
        }

        internal void Apply(Control primary, string text, bool showHint, Control anchorOverride, params Control[] fallbackTargets)
        {            if (primary == null)
            {
                return;
            }

            Control[] normalizedFallbackTargets = fallbackTargets ?? new Control[0];
            _anchorOverridesByPrimary[primary] = anchorOverride;
            _fallbackTargetsByPrimary[primary] = normalizedFallbackTargets;
            _showHintByPrimary[primary] = showHint;
            Track(primary);            if (anchorOverride != null)
            {
                Track(anchorOverride);
            }
            if (normalizedFallbackTargets.Length > 0)
            {
                foreach (Control target in normalizedFallbackTargets)
                {
                    Track(target);
                }
            }

            _toolTip.SetToolTip(primary, text);
            if (normalizedFallbackTargets.Length > 0)
            {
                foreach (Control target in normalizedFallbackTargets)
                {                    if (target != null && !ReferenceEquals(target, primary))
                    {
                        _toolTip.SetToolTip(target, text);
                    }
                }
            }

            UpdateHint(primary, text, showHint, anchorOverride, normalizedFallbackTargets);
        }

        private void UpdateHint(Control primary, string text, bool showHint, Control anchorOverride, Control[] fallbackTargets)
        {            if (primary == null)
            {
                return;
            }

            Control anchor = ResolveAnchor(primary, anchorOverride, fallbackTargets);
            bool shouldShow = showHint
                              && !string.IsNullOrWhiteSpace(text)
                              && primary.Visible
                              && anchor != null
                              && anchor.Visible
                              && anchor.Parent != null;
            if (!shouldShow)
            {
                HideHint(primary);
                return;
            }

            Label hint = GetOrCreateHint(primary);
            if (!ReferenceEquals(hint.Parent, anchor.Parent))
            {                if (hint.Parent != null)
                {
                    hint.Parent.Controls.Remove(hint);
                }

                anchor.Parent.Controls.Add(hint);
            }

            hint.Visible = true;
            _toolTip.SetToolTip(hint, text);
            hint.AccessibleName = string.IsNullOrWhiteSpace(text) ? "?" : text;
            hint.AccessibleDescription = string.IsNullOrWhiteSpace(text) ? string.Empty : text;
            PositionHint(anchor, hint);
            hint.BringToFront();
        }

        private Label GetOrCreateHint(Control primary)
        {
            Label hint;
            if (_hintLabels.TryGetValue(primary, out hint))
            {
                return hint;
            }

            hint = new Label
            {
                AutoSize = false,
                Size = new Size(HintSize, HintSize),
                Text = "?",
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Help,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.FromArgb(240, 244, 248),
                ForeColor = Color.FromArgb(0, 102, 153),
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 8f, FontStyle.Bold, GraphicsUnit.Point),
                TabStop = false,
                Visible = false
            };

            _hintLabels[primary] = hint;
            Track(primary);
            return hint;
        }

        private void Track(Control primary)
        {            if (primary == null || _trackedControls.Contains(primary))
            {
                return;
            }

            _trackedControls.Add(primary);
            primary.EnabledChanged += OnTrackedControlChanged;
            primary.VisibleChanged += OnTrackedControlChanged;
            primary.LocationChanged += OnTrackedControlChanged;
            primary.SizeChanged += OnTrackedControlChanged;
            primary.ParentChanged += OnTrackedControlChanged;
        }

        private void OnTrackedControlChanged(object sender, EventArgs e)
        {
            Control changed = sender as Control;            if (changed == null)
            {
                return;
            }

            RefreshHint(changed);

            foreach (KeyValuePair<Control, Control[]> pair in _fallbackTargetsByPrimary)
            {
                Control primary = pair.Key;
                if (ReferenceEquals(primary, changed))
                {
                    continue;
                }

                Control anchorOverride;
                if (_anchorOverridesByPrimary.TryGetValue(primary, out anchorOverride)
                    && ReferenceEquals(anchorOverride, changed))
                {
                    RefreshHint(primary);
                    continue;
                }

                Control[] fallbackTargets = pair.Value ?? new Control[0];
                for (int i = 0; i < fallbackTargets.Length; i++)
                {
                    if (ReferenceEquals(fallbackTargets[i], changed))
                    {
                        RefreshHint(primary);
                        break;
                    }
                }
            }
        }

        private void RefreshHint(Control primary)
        {            if (primary == null)
            {
                return;
            }

            Control[] fallbackTargets;
            if (!_fallbackTargetsByPrimary.TryGetValue(primary, out fallbackTargets))
            {
                fallbackTargets = new Control[0];
            }

            bool showHint;
            if (!_showHintByPrimary.TryGetValue(primary, out showHint))
            {
                showHint = false;
            }

            Control anchorOverride;
            if (!_anchorOverridesByPrimary.TryGetValue(primary, out anchorOverride))
            {
                anchorOverride = null;
            }

            UpdateHint(primary, _toolTip.GetToolTip(primary), showHint, anchorOverride, fallbackTargets);
        }

        private static Control ResolveAnchor(Control primary, Control anchorOverride, Control[] fallbackTargets)
        {            if (primary == null)
            {
                return null;
            }            if (anchorOverride != null && anchorOverride.Visible && anchorOverride.Parent != null)
            {
                return anchorOverride;
            }

            if (primary is CheckBox || primary is RadioButton || primary is Label || primary is LinkLabel)
            {
                return primary;
            }            if (fallbackTargets != null)
            {
                foreach (Control target in fallbackTargets)
                {                    if (target == null || !target.Visible)
                    {
                        continue;
                    }

                    if (target is Label || target is LinkLabel || target is CheckBox || target is RadioButton)
                    {
                        return target;
                    }
                }
            }

            return primary;
        }

        private static void PositionHint(Control anchor, Control hint)
        {
            Control parent = anchor.Parent;            if (parent == null)
            {
                return;
            }

            int x = anchor.Right + HintSpacing;
            int y = anchor.Top + Math.Max(0, (anchor.Height - hint.Height) / 2);
            int maxX = Math.Max(HintSpacing, parent.ClientSize.Width - hint.Width - HintSpacing);
            int rightNeighborLeft = FindNearestRightNeighborLeft(anchor, hint, parent, y, hint.Height);
            if (rightNeighborLeft > 0)
            {
                maxX = Math.Min(maxX, rightNeighborLeft - hint.Width - HintSpacing);
            }

            if (x > maxX)
            {
                x = maxX;
            }

            if (x < HintSpacing)
            {
                x = HintSpacing;
            }

            int maxY = Math.Max(0, parent.ClientSize.Height - hint.Height - HintSpacing);
            hint.Location = new Point(x, Math.Min(y, maxY));
        }

        private static int FindNearestRightNeighborLeft(Control anchor, Control hint, Control parent, int hintTop, int hintHeight)
        {            if (anchor == null || parent == null)
            {
                return 0;
            }

            int nearestLeft = int.MaxValue;
            int hintBottom = hintTop + hintHeight;

            foreach (Control sibling in parent.Controls)
            {                if (sibling == null || !sibling.Visible || ReferenceEquals(sibling, anchor) || ReferenceEquals(sibling, hint))
                {
                    continue;
                }

                if (sibling.Left <= anchor.Right)
                {
                    continue;
                }

                bool overlapsVertically = sibling.Bottom > hintTop && sibling.Top < hintBottom;
                if (!overlapsVertically)
                {
                    continue;
                }

                if (sibling.Left < nearestLeft)
                {
                    nearestLeft = sibling.Left;
                }
            }

            return nearestLeft == int.MaxValue ? 0 : nearestLeft;
        }

        private void HideHint(Control primary)
        {
            Label hint;
            if (_hintLabels.TryGetValue(primary, out hint))
            {
                hint.Visible = false;
                _toolTip.SetToolTip(hint, string.Empty);
            }
        }
    }
}

