using Toybox.WatchUi;
using Toybox.Graphics;
using Toybox.Timer;

const COLOR_BG = 0x000000;
const COLOR_TEXT = 0xFFFFFF;
const COLOR_MUTED = 0x808080;
const COLOR_TRACK = 0x333333;
const COLOR_GOOD = 0x3CD682;
const COLOR_WARN = 0xFFAA00;
const COLOR_OVER = 0xFF4433;

// The one screen: a ring for how much of today's allowance is gone, a big
// number for how many sets are still on the table, and the raw numbers below.
class CounterView extends WatchUi.View {

    hidden var _model;
    hidden var _sync;
    hidden var _timer;

    function initialize(model, sync) {
        View.initialize();
        _model = model;
        _sync = sync;
    }

    function onShow() {
        _model.rolloverIfNeeded();
        // Active calories drift up on their own, so repaint on a slow tick.
        _timer = new Timer.Timer();
        _timer.start(method(:onTick), 30000, true);
    }

    function onHide() {
        if (_timer != null) {
            _timer.stop();
            _timer = null;
        }
    }

    function onTick() {
        if (_model.rolloverIfNeeded()) {
            _sync.syncNow(true);
        }
        WatchUi.requestUpdate();
    }

    function onUpdate(dc) {
        var width = dc.getWidth();
        var height = dc.getHeight();
        var centerX = width / 2;
        var centerY = height / 2;

        dc.setColor(COLOR_TEXT, COLOR_BG);
        dc.clear();

        var fraction = _model.consumedFraction();
        var accent = accentColor(fraction);

        drawRing(dc, centerX, centerY, width, height, fraction, accent);

        var center = Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER;

        // Sets logged, small, above the hero number.
        var setsLine = _model.clicks().toString() + " SETS  x" + _model.kcalPerClick.toString();
        dc.setColor(COLOR_MUTED, Graphics.COLOR_TRANSPARENT);
        dc.drawText(centerX, px(height, 0.21), Graphics.FONT_XTINY, setsLine, center);

        // Hero: sets still available today. Past the budget it counts the other
        // way, so the big number stays positive and the label carries the sign.
        var left = _model.remainingUnits();
        var hero = (left >= 0) ? left : -left;
        dc.setColor(accent, Graphics.COLOR_TRANSPARENT);
        dc.drawText(centerX, px(height, 0.41), heroFont(), hero.toString(), center);

        dc.setColor(COLOR_MUTED, Graphics.COLOR_TRANSPARENT);
        dc.drawText(centerX, px(height, 0.575), Graphics.FONT_XTINY,
            (left >= 0) ? "SETS LEFT" : "SETS OVER", center);

        // Separator.
        dc.setColor(COLOR_TRACK, Graphics.COLOR_TRANSPARENT);
        dc.setPenWidth(1);
        dc.drawLine(px(width, 0.30), px(height, 0.635), px(width, 0.70), px(height, 0.635));

        // Eaten out of the day's allowance.
        dc.setColor(COLOR_TEXT, Graphics.COLOR_TRANSPARENT);
        dc.drawText(centerX, px(height, 0.705), Graphics.FONT_TINY,
            _model.consumedKcal().toString() + " / " + _model.allowanceKcal().toString(), center);

        // Where the extra room came from.
        var activeLine = "BASE " + _model.dailyBudget.toString();
        if (_model.includeActive) {
            activeLine = activeLine + "  +" + _model.activeCalories().toString() + " ACTIVE";
        }
        dc.setColor(COLOR_MUTED, Graphics.COLOR_TRANSPARENT);
        dc.drawText(centerX, px(height, 0.785), Graphics.FONT_XTINY, activeLine, center);

        var syncLine = _sync.statusLabel();
        if (syncLine.length() > 0) {
            dc.setColor(syncColor(), Graphics.COLOR_TRANSPARENT);
            dc.drawText(centerX, px(height, 0.855), Graphics.FONT_XTINY, syncLine, center);
        }
    }

    // Pixel position from a fraction of the screen, as the integer the Dc wants.
    hidden function px(extent, fraction) {
        return (extent * fraction).toNumber();
    }

    hidden function drawRing(dc, centerX, centerY, width, height, fraction, accent) {
        var shorter = width < height ? width : height;
        var pen = (shorter * 0.045).toNumber();
        if (pen < 3) {
            pen = 3;
        }
        var radius = (shorter / 2) - (pen / 2) - 2;

        dc.setPenWidth(pen);
        dc.setColor(COLOR_TRACK, Graphics.COLOR_TRANSPARENT);
        dc.drawCircle(centerX, centerY, radius);

        if (fraction <= 0.005) {
            return;
        }

        var filled = fraction > 1.0 ? 1.0 : fraction;
        dc.setColor(accent, Graphics.COLOR_TRANSPARENT);
        if (filled >= 0.995) {
            dc.drawCircle(centerX, centerY, radius);
            return;
        }

        // Degrees run counter-clockwise from 3 o'clock, so sweeping clockwise
        // from the top (90) means subtracting.
        var end = 90.0 - (360.0 * filled);
        while (end < 0.0) {
            end = end + 360.0;
        }
        dc.drawArc(centerX, centerY, radius, Graphics.ARC_CLOCKWISE, 90, end.toNumber());
    }

    hidden function accentColor(fraction) {
        if (fraction >= 1.0) {
            return COLOR_OVER;
        }
        if (fraction >= 0.85) {
            return COLOR_WARN;
        }
        return COLOR_GOOD;
    }

    hidden function syncColor() {
        if (_sync.status == :failed) {
            return COLOR_OVER;
        }
        if (_sync.status == :offline) {
            return COLOR_WARN;
        }
        return COLOR_MUTED;
    }

    // The tall number fonts are not on every device; fall back gracefully.
    hidden function heroFont() {
        if (Graphics has :FONT_NUMBER_THAI_HOT) {
            return Graphics.FONT_NUMBER_THAI_HOT;
        }
        if (Graphics has :FONT_NUMBER_HOT) {
            return Graphics.FONT_NUMBER_HOT;
        }
        return Graphics.FONT_LARGE;
    }
}
