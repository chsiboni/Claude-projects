using Toybox.Application;
using Toybox.WatchUi;

class CalorieApp extends Application.AppBase {

    hidden var _model;
    hidden var _sync;

    function initialize() {
        AppBase.initialize();
    }

    function onStart(state) {
        _model = new CounterModel();
        _model.load();
        _sync = new SyncManager(_model);
        // Catch up on anything the last session could not deliver.
        _sync.syncNow(false);
    }

    function onStop(state) {
        if (_model != null) {
            _model.save();
        }
    }

    function getInitialView() {
        return [ new CounterView(_model, _sync), new CounterDelegate(_model, _sync) ];
    }

    function onSettingsChanged() {
        if (_model != null) {
            _model.reloadSettings();
        }
        WatchUi.requestUpdate();
    }
}
