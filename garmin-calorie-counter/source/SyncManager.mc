using Toybox.Communications;
using Toybox.System;
using Toybox.WatchUi;

// Pushes the day's totals to the nutrition server. Fire-and-forget: if a POST
// fails the counter stays marked dirty and the next attempt sends it again.
class SyncManager {

    hidden var _model;
    hidden var _busy;
    var status; // :off :idle :syncing :ok :failed :offline

    function initialize(model) {
        _model = model;
        _busy = false;
        status = :idle;
    }

    function configured() {
        var url = _model.serverUrl;
        return (url != null) && (url.length() > 0);
    }

    function busy() {
        return _busy;
    }

    // force=true sends even when nothing changed (manual "Sync now").
    function syncNow(force) {
        if (!configured()) {
            status = :off;
            return;
        }
        if (_busy) {
            return;
        }
        if (!force && !_model.pendingSync()) {
            status = :ok;
            return;
        }

        var settings = System.getDeviceSettings();
        if ((settings has :phoneConnected) && (settings.phoneConnected != true)) {
            status = :offline;
            return;
        }

        var headers = { "Content-Type" => Communications.REQUEST_CONTENT_TYPE_JSON };
        var key = _model.apiKey;
        if ((key != null) && (key.length() > 0)) {
            headers.put("X-Api-Key", key);
        }

        var options = {
            :method => Communications.HTTP_REQUEST_METHOD_POST,
            :headers => headers,
            :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON
        };

        _busy = true;
        status = :syncing;
        Communications.makeWebRequest(_model.serverUrl, _model.payload(), options, method(:onResponse));
    }

    function onResponse(responseCode, data) {
        _busy = false;
        if (responseCode >= 200 && responseCode < 300) {
            status = :ok;
            _model.markSynced();
        } else if (responseCode == Communications.BLE_CONNECTION_UNAVAILABLE ||
                   responseCode == Communications.BLE_HOST_TIMEOUT) {
            status = :offline;
        } else {
            status = :failed;
        }
        WatchUi.requestUpdate();
    }

    // Short label for the bottom of the main screen.
    function statusLabel() {
        if (!configured()) {
            return "";
        }
        if (status == :syncing) {
            return "SYNCING";
        }
        if (status == :offline) {
            return "NO PHONE";
        }
        if (status == :failed) {
            return "SYNC FAILED";
        }
        if (_model.pendingSync()) {
            return "PENDING";
        }
        if (status == :ok) {
            return "SYNCED";
        }
        return "";
    }
}
