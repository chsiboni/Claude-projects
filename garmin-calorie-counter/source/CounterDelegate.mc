using Toybox.WatchUi;
using Toybox.System;
using Toybox.Attention;

// START / tap  -> log a set
// DOWN         -> take one back
// UP           -> sync now
// MENU (hold)  -> menu
class CounterDelegate extends WatchUi.BehaviorDelegate {

    hidden var _model;
    hidden var _sync;

    function initialize(model, sync) {
        BehaviorDelegate.initialize();
        _model = model;
        _sync = sync;
    }

    function onKey(keyEvent) {
        var key = keyEvent.getKey();
        if (key == WatchUi.KEY_ENTER) {
            return count(1);
        }
        if (key == WatchUi.KEY_DOWN) {
            return count(-1);
        }
        if (key == WatchUi.KEY_UP) {
            _sync.syncNow(true);
            WatchUi.requestUpdate();
            return true;
        }
        return false;
    }

    // Screen tap on touch models, and the select behaviour generally.
    function onSelect() {
        return count(1);
    }

    function onMenu() {
        WatchUi.pushView(buildMenu(), new MainMenuDelegate(_model, _sync), WatchUi.SLIDE_UP);
        return true;
    }

    function onBack() {
        _model.save();
        _sync.syncNow(false);
        return false; // let the system close the app
    }

    hidden function count(delta) {
        var moved = _model.add(delta);
        if (moved) {
            feedback(delta > 0);
            if (_model.autoSync) {
                _sync.syncNow(false);
            }
        }
        WatchUi.requestUpdate();
        return true;
    }

    hidden function feedback(added) {
        if (!(Toybox has :Attention)) {
            return;
        }
        var settings = System.getDeviceSettings();
        if ((Attention has :vibrate) && (settings.vibrateOn == true)) {
            Attention.vibrate([new Attention.VibeProfile(50, (added ? 60 : 30))]);
        }
        if ((Attention has :playTone) && (settings.tonesOn == true)) {
            Attention.playTone((added ? Attention.TONE_KEY : Attention.TONE_RESET));
        }
    }

    hidden function buildMenu() {
        var menu = new WatchUi.Menu2({ :title => WatchUi.loadResource(Rez.Strings.MenuTitle) });
        menu.addItem(new WatchUi.MenuItem(WatchUi.loadResource(Rez.Strings.MenuUndo), null, :undo, {}));
        menu.addItem(new WatchUi.MenuItem(WatchUi.loadResource(Rez.Strings.MenuSync), null, :sync, {}));
        menu.addItem(new WatchUi.MenuItem(WatchUi.loadResource(Rez.Strings.MenuReset), null, :reset, {}));
        return menu;
    }
}
