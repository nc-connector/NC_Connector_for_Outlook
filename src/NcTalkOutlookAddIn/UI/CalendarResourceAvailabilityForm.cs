// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    internal sealed class CalendarResourceAvailabilityForm : ScaledForm
    {
        private readonly CalendarResource _resource;
        private readonly CalDavFreeBusyService _service;
        private readonly DateTimePicker _date = new DateTimePicker { Format = DateTimePickerFormat.Short };
        private readonly Label _zoneLabel = new Label { AutoSize = true };
        private readonly Label _status = new Label { AutoSize = true };
        private readonly DayTimeline _timeline = new DayTimeline { Dock = DockStyle.Fill };
        private readonly Button _previous = new Button();
        private readonly Button _next = new Button();
        private CancellationTokenSource _cancellation;
        private TimeZoneInfo _timeZone;
        private DateTime _appointmentStart, _appointmentEnd;
        private int _generation;
        private bool _shown;

        internal CalendarResourceAvailabilityForm(CalendarResource resource, CalDavFreeBusyService service,
            DateTime appointmentStartUtc, DateTime appointmentEndUtc, TimeZoneInfo timeZone)
        {
            _resource = resource;
            _service = service;
            _timeZone = timeZone;
            _appointmentStart = appointmentStartUtc;
            _appointmentEnd = appointmentEndUtc;
            Text = Strings.ResourceAvailabilityTitle + " — " + resource.DisplayName;
            Icon = BrandingAssets.GetAppIcon(32);
            Font = new Font("Segoe UI", 9F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(500, 740);
            MinimumSize = new Size(440, 580);
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(14), ColumnCount = 1, RowCount = 6 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int row = 0; row < 6; row++) { layout.RowStyles.Add(new RowStyle(row == 4 ? SizeType.Percent : SizeType.AutoSize, row == 4 ? 100 : 0)); }
            layout.Controls.Add(new Label { Text = resource.DisplayName, AutoSize = true, UseMnemonic = false, Font = new Font(Font, FontStyle.Bold) }, 0, 0);
            var navigation = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top, WrapContents = true };
            _previous.Text = "‹";
            _previous.Width = 34;
            _previous.AccessibleName = Strings.ResourcePreviousDay;
            _next.Text = "›";
            _next.Width = 34;
            _next.AccessibleName = Strings.ResourceNextDay;
            _date.Width = 135;
            _date.AccessibleName = Strings.ResourceDateLabel;
            _date.Value = TimeZoneInfo.ConvertTimeFromUtc(appointmentStartUtc, timeZone).Date;
            var today = new Button { Text = Strings.ResourceToday, AutoSize = true };
            _previous.Click += delegate { ChangeDay(-1); };
            _next.Click += delegate { ChangeDay(1); };
            today.Click += delegate { _date.Value = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, _timeZone).Date; };
            _date.ValueChanged += async delegate { if (_shown) { await LoadDayAsync(); } };
            navigation.Controls.Add(_previous);
            navigation.Controls.Add(_date);
            navigation.Controls.Add(_next);
            navigation.Controls.Add(today);
            layout.Controls.Add(navigation, 0, 1);
            layout.Controls.Add(_zoneLabel, 0, 2);
            layout.Controls.Add(_status, 0, 3);
            layout.Controls.Add(_timeline, 0, 4);
            var close = new Button { Text = Strings.ResourceCloseButton, DialogResult = DialogResult.Cancel, AutoSize = true, Anchor = AnchorStyles.Right, MinimumSize = new Size(100, 30) };
            layout.Controls.Add(close, 0, 5);
            Controls.Add(layout);
            CancelButton = close;
            KeyPreview = true;
            KeyDown += (s, e) =>
            {
                if (e.Alt && (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right))
                { ChangeDay(e.KeyCode == Keys.Left ? -1 : 1); e.Handled = true; e.SuppressKeyPress = true; }
            };
            UiThemeManager.ApplyToForm(this);
            UpdateLabels();
        }

        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _shown = true;
            Rectangle area = Screen.FromControl(this).WorkingArea;
            if (Height > area.Height) { Height = area.Height; Top = area.Top; }
            await LoadDayAsync();
        }

        internal async Task UpdateAppointmentWindow(DateTime startUtc, DateTime endUtc, TimeZoneInfo zone)
        {
            bool changedZone = _timeZone.Id != zone.Id;
            _appointmentStart = startUtc;
            _appointmentEnd = endUtc;
            _timeZone = zone;
            UpdateLabels();
            _timeline.AppointmentStart = startUtc;
            _timeline.AppointmentEnd = endUtc;
            _timeline.Invalidate();
            if (changedZone && _shown) { await LoadDayAsync(); }
        }

        private void UpdateLabels()
        {
            _zoneLabel.Text = Strings.ResourceTimeZoneLabel + ": " + _timeZone.DisplayName;
            _timeline.AccessibleName = Strings.ResourceAvailabilityTitle;
        }

        private void ChangeDay(int days)
        {
            DateTime next = _date.Value.Date.AddDays(days);
            if (next >= _date.MinDate && next <= _date.MaxDate) { _date.Value = next; }
        }

        internal static DateTime DayBoundaryUtc(DateTime date, TimeZoneInfo zone)
        {
            DateTime local = DateTime.SpecifyKind(date.Date, DateTimeKind.Unspecified);
            // Some zones move their clocks at midnight; use the first actual instant of that date.
            for (int minute = 0; minute < 1440 && zone.IsInvalidTime(local); minute++) { local = local.AddMinutes(1); }
            if (local.Date != date.Date) { throw new ArgumentException("The selected local date does not exist in this time zone."); }
            if (zone.IsAmbiguousTime(local))
            { return DateTime.SpecifyKind(local - zone.GetAmbiguousTimeOffsets(local).Max(), DateTimeKind.Utc); }
            return TimeZoneInfo.ConvertTimeToUtc(local, zone);
        }

        private async Task LoadDayAsync()
        {
            if (_cancellation != null) { _cancellation.Cancel(); }
            var cancellation = new CancellationTokenSource();
            _cancellation = cancellation;
            int generation = ++_generation;
            _status.Text = Strings.ResourceChecking;
            _timeline.Data = null;
            _timeline.AccessibleDescription = Strings.ResourceChecking;
            _timeline.TimeZone = _timeZone;
            _timeline.AppointmentStart = _appointmentStart;
            _timeline.AppointmentEnd = _appointmentEnd;
            _previous.Enabled = _date.Value.Date > _date.MinDate.Date;
            _next.Enabled = _date.Value.Date < _date.MaxDate.Date;
            try
            {
                DateTime start = DayBoundaryUtc(_date.Value.Date, _timeZone);
                DateTime end = DayBoundaryUtc(_date.Value.Date.AddDays(1), _timeZone);
                _timeline.StartUtc = start;
                _timeline.EndUtc = end;
                _timeline.Invalidate();
                var results = await _service.QueryAsync(new[] { _resource }, start, end, cancellation.Token);
                if (IsDisposed || cancellation.IsCancellationRequested || generation != _generation) { return; }
                CalendarResourceAvailability data;
                _timeline.Data = results.TryGetValue(_resource.Email, out data) ? data : null;
                _status.Text = data != null && data.IsKnown ? "● " + Strings.ResourceAvailable + "    ● " + Strings.ResourceBusy + "    ▯ " + Strings.ResourceAppointmentIntervalLabel : Strings.ResourceUnknown;
                _timeline.AccessibleDescription = data != null && data.IsKnown
                    ? (data.BusyPeriods.Count == 0 ? Strings.ResourceAvailable : string.Join("; ", data.BusyPeriods.Select(p => Strings.ResourceBusy + " " + TimeZoneInfo.ConvertTimeFromUtc(p.StartUtc, _timeZone).ToString("t") + "–" + TimeZoneInfo.ConvertTimeFromUtc(p.EndUtc, _timeZone).ToString("t"))))
                    : Strings.ResourceUnknown;
            }
            catch (OperationCanceledException) { return; }
            catch (Exception ex)
            {
                if (IsDisposed || generation != _generation) { return; }
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Resource day availability could not be loaded.", ex);
                _timeline.Data = null;
                _timeline.AccessibleDescription = Strings.ResourceUnknown;
                _status.Text = Strings.ResourceAvailabilityFailed + " " + Strings.ResourceUnknown;
            }
            finally
            {
                if (ReferenceEquals(_cancellation, cancellation)) { _cancellation = null; }
                cancellation.Dispose();
            }
            if (!IsDisposed && generation == _generation) { _timeline.Invalidate(); }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { _generation++; if (_cancellation != null) { _cancellation.Cancel(); } }
            base.Dispose(disposing);
        }

        private sealed class DayTimeline : Control
        {
            internal DateTime StartUtc, EndUtc, AppointmentStart, AppointmentEnd;
            internal TimeZoneInfo TimeZone;
            internal CalendarResourceAvailability Data;
            internal DayTimeline()
            {
                DoubleBuffered = true;
                TabStop = true;
                AccessibleRole = AccessibleRole.Chart;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                if (EndUtc <= StartUtc || TimeZone == null) { return; }
                int labelWidth = Math.Max(90, DeviceDpi * 100 / 96);
                Rectangle plot = new Rectangle(labelWidth, 10, Math.Max(1, Width - labelWidth - 10), Math.Max(1, Height - 30));
                bool known = Data != null && Data.IsKnown;
                Color free = SystemInformation.HighContrast ? SystemColors.Window : Color.FromArgb(44, 115, 66);
                Color busy = SystemInformation.HighContrast ? SystemColors.Highlight : Color.FromArgb(185, 51, 55);
                using (var background = new SolidBrush(known ? free : SystemColors.ControlDark)) { e.Graphics.FillRectangle(background, plot); }
                if (known)
                {
                    using (var fill = new SolidBrush(busy))
                    foreach (CalendarBusyPeriod period in Data.BusyPeriods)
                    {
                        float top = Position(period.StartUtc, plot);
                        float bottom = Position(period.EndUtc, plot);
                        e.Graphics.FillRectangle(fill, plot.Left, top, plot.Width, Math.Max(2, bottom - top));
                        if (bottom - top >= Font.Height)
                        { TextRenderer.DrawText(e.Graphics, Strings.ResourceBusy, Font, new Rectangle(plot.Left + 4, (int)top, plot.Width - 8, (int)(bottom - top)), Color.White, TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis); }
                    }
                }
                using (var line = new Pen(Color.FromArgb(90, ForeColor)))
                {
                    for (DateTime hour = StartUtc; hour <= EndUtc; hour = hour.AddHours(1))
                    {
                        int y = (int)Position(hour, plot);
                        DateTime local = TimeZoneInfo.ConvertTimeFromUtc(hour, TimeZone);
                        string label = local.ToString("HH:mm", CultureInfo.CurrentCulture);
                        if (TimeZone.IsAmbiguousTime(local)) { label += " " + new DateTimeOffset(hour).ToOffset(TimeZone.GetUtcOffset(hour)).ToString("zzz"); }
                        TextRenderer.DrawText(e.Graphics, label, Font, new Rectangle(0, y - Font.Height / 2, labelWidth - 6, Font.Height + 2), ForeColor, TextFormatFlags.Right | TextFormatFlags.NoPadding);
                        e.Graphics.DrawLine(line, plot.Left, y, plot.Right, y);
                    }
                }
                if (AppointmentStart < EndUtc && AppointmentEnd > StartUtc)
                {
                    float top = Position(AppointmentStart < StartUtc ? StartUtc : AppointmentStart, plot);
                    float bottom = Position(AppointmentEnd > EndUtc ? EndUtc : AppointmentEnd, plot);
                    using (var pen = new Pen(SystemInformation.HighContrast ? SystemColors.WindowText : Color.FromArgb(80, 185, 255), 3))
                    { e.Graphics.DrawRectangle(pen, plot.Left + 2, top + 1, plot.Width - 4, Math.Max(3, bottom - top - 2)); }
                }
                if (Focused) { ControlPaint.DrawFocusRectangle(e.Graphics, ClientRectangle, ForeColor, BackColor); }
            }

            private float Position(DateTime time, Rectangle plot)
            { return plot.Top + (float)((time - StartUtc).TotalSeconds / (EndUtc - StartUtc).TotalSeconds * plot.Height); }
        }
    }
}
