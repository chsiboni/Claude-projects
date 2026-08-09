using Toybox.WatchUi;

class MainMenuDelegate extends WatchUi.Menu2InputDelegate {

    hidden var _model;
    hidden var _sync;

    function initialize(model, sync) {
        Menu2InputDelegate.initialize();
        _model = model;
        _sync = sync;
    }

    function onSelect(item) {
        var id = item.getId();

        if (id == :undo) {
            _model.add(-1);
            if (_model.autoSync) {
                _sync.syncNow(false);
            }
            WatchUi.popView(WatchUi.SLIDE_DOWN);

        } else if (id == :sync) {
            _sync.syncNow(true);
            WatchUi.popView(WatchUi.SLIDE_DOWN);

        } else if (id == :reset) {
            var prompt = new WatchUi.Confirmation(WatchUi.loadResource(Rez.Strings.MenuResetConfirm));
            WatchUi.pushView(prompt, new ResetConfirmDelegate(_model, _sync), WatchUi.SLIDE_LEFT);
        }
    }

    function onBack() {
        WatchUi.popView(WatchUi.SLIDE_DOWN);
    }
}

class ResetConfirmDelegate extends WatchUi.ConfirmationDelegate {

    hidden var _model;
    hidden var _sync;

    function initialize(model, sync) {
        ConfirmationDelegate.initialize();
        _model = model;
        _sync = sync;
    }

    function onResponse(response) {
        if (response == WatchUi.CONFIRM_YES) {
            _model.reset();
            _sync.syncNow(true);
            WatchUi.requestUpdate();
        }
        return true;
    }
}
