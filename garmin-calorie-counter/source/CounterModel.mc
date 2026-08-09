using Toybox.Application;
using Toybox.ActivityMonitor;
using Toybox.UserProfile;
using Toybox.System;
using Toybox.Time;
using Toybox.Time.Gregorian;
using Toybox.Math;

const STORE_CLICKS = "clicks";
const STORE_DAY = "day";
const STORE_SYNCED = "syncedClicks";

// Holds the day's state: how many "sets" were eaten, what a set is worth, and
// how much room the day's activity has opened up.
class CounterModel {

    hidden var _clicks;
    hidden var _day;
    hidden var _syncedClicks;

    var kcalPerClick;
    var dailyBudget;
    var includeActive;
    var autoSync;
    var serverUrl;
    var apiKey;

    function initialize() {
        _clicks = 0;
        _day = null;
        _syncedClicks = -1;
        reloadSettings();
    }

    // ---- settings -------------------------------------------------------

    function reloadSettings() {
        kcalPerClick = numberSetting("kcalPerClick", 50, 1, 2000);
        dailyBudget = numberSetting("dailyBudget", 2000, 0, 20000);
        includeActive = boolSetting("includeActive", true);
        autoSync = boolSetting("autoSync", true);
        serverUrl = stringSetting("serverUrl");
        apiKey = stringSetting("apiKey");
    }

    hidden function rawProperty(key) {
        var value = null;
        try {
            value = Application.Properties.getValue(key);
        } catch (ex) {
            value = null;
        }
        return value;
    }

    hidden function numberSetting(key, fallback, min, max) {
        var value = rawProperty(key);
        if (value == null) {
            return fallback;
        }
        var number = fallback;
        try {
            number = value.toNumber();
        } catch (ex) {
            return fallback;
        }
        if (number == null || number < min || number > max) {
            return fallback;
        }
        return number;
    }

    hidden function boolSetting(key, fallback) {
        var value = rawProperty(key);
        if (value == null) {
            return fallback;
        }
        return (value == true);
    }

    hidden function stringSetting(key) {
        var value = rawProperty(key);
        if (value == null) {
            return "";
        }
        return value.toString();
    }

    // ---- persistence ----------------------------------------------------

    hidden function stored(key, fallback) {
        var value = null;
        try {
            value = Application.Storage.getValue(key);
        } catch (ex) {
            value = null;
        }
        if (value == null) {
            return fallback;
        }
        return value;
    }

    function load() {
        reloadSettings();
        _day = stored(STORE_DAY, null);
        _clicks = stored(STORE_CLICKS, 0);
        _syncedClicks = stored(STORE_SYNCED, -1);
        rolloverIfNeeded();
    }

    function save() {
        try {
            Application.Storage.setValue(STORE_CLICKS, _clicks);
            Application.Storage.setValue(STORE_DAY, _day);
            Application.Storage.setValue(STORE_SYNCED, _syncedClicks);
        } catch (ex) {
            // Storage full or unavailable - the in-memory value still stands.
        }
    }

    // A new calendar day starts from zero. Returns true when it just rolled.
    function rolloverIfNeeded() {
        var today = todayKey();
        if (_day != null && _day.equals(today)) {
            return false;
        }
        _day = today;
        _clicks = 0;
        _syncedClicks = -1;
        save();
        return true;
    }

    function todayKey() {
        var now = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
        return now.year.format("%04d") + "-" + now.month.format("%02d") + "-" + now.day.format("%02d");
    }

    // ---- counting -------------------------------------------------------

    function clicks() {
        return _clicks;
    }

    function day() {
        return _day;
    }

    // Adds (or removes) sets. Never drops below zero. Returns true if it moved.
    function add(delta) {
        rolloverIfNeeded();
        var next = _clicks + delta;
        if (next < 0) {
            next = 0;
        }
        if (next == _clicks) {
            return false;
        }
        _clicks = next;
        save();
        return true;
    }

    function reset() {
        rolloverIfNeeded();
        _clicks = 0;
        save();
    }

    // ---- calories -------------------------------------------------------

    // Calories burned by movement today, on top of resting metabolism.
    function activeCalories() {
        var info = null;
        try {
            info = ActivityMonitor.getInfo();
        } catch (ex) {
            info = null;
        }
        if (info == null) {
            return 0;
        }

        // Newer devices expose this directly.
        if ((info has :activeCalories) && (info.activeCalories != null)) {
            return info.activeCalories;
        }

        // Otherwise subtract the resting burn from the day's total.
        if (!(info has :calories) || info.calories == null) {
            return 0;
        }
        var active = info.calories - restingBurnSoFar();
        if (active < 0) {
            return 0;
        }
        return active;
    }

    function totalCalories() {
        var info = null;
        try {
            info = ActivityMonitor.getInfo();
        } catch (ex) {
            info = null;
        }
        if (info == null || !(info has :calories) || info.calories == null) {
            return 0;
        }
        return info.calories;
    }

    function steps() {
        var info = null;
        try {
            info = ActivityMonitor.getInfo();
        } catch (ex) {
            info = null;
        }
        if (info == null || !(info has :steps) || info.steps == null) {
            return 0;
        }
        return info.steps;
    }

    // Mifflin-St Jeor resting rate, scaled to the part of the day elapsed.
    hidden function restingBurnSoFar() {
        var profile = null;
        try {
            profile = UserProfile.getProfile();
        } catch (ex) {
            profile = null;
        }

        var kg = 75.0;
        var cm = 175.0;
        var age = 35.0;
        var female = false;

        if (profile != null) {
            if ((profile has :weight) && (profile.weight != null) && (profile.weight > 0)) {
                kg = profile.weight / 1000.0;
            }
            if ((profile has :height) && (profile.height != null) && (profile.height > 0)) {
                cm = profile.height * 1.0;
            }
            if ((profile has :birthYear) && (profile.birthYear != null)) {
                var born = profile.birthYear;
                if (born < 1900) {
                    born = born + 1900;
                }
                var years = Gregorian.info(Time.now(), Time.FORMAT_SHORT).year - born;
                if (years > 5 && years < 110) {
                    age = years * 1.0;
                }
            }
            if ((profile has :gender) && (profile.gender != null) && (profile.gender == UserProfile.GENDER_FEMALE)) {
                female = true;
            }
        }

        var bmr = 10.0 * kg + 6.25 * cm - 5.0 * age;
        bmr = female ? bmr - 161.0 : bmr + 5.0;

        var now = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
        var elapsed = (now.hour * 3600 + now.min * 60 + now.sec) / 86400.0;
        return (bmr * elapsed).toNumber();
    }

    function consumedKcal() {
        return _clicks * kcalPerClick;
    }

    function allowanceKcal() {
        var allowance = dailyBudget;
        if (includeActive) {
            allowance = allowance + activeCalories();
        }
        return allowance;
    }

    function remainingKcal() {
        return allowanceKcal() - consumedKcal();
    }

    // How many more sets fit in today's budget. Negative means over.
    function remainingUnits() {
        if (kcalPerClick <= 0) {
            return 0;
        }
        return Math.floor(remainingKcal().toFloat() / kcalPerClick).toNumber();
    }

    // 0.0 to 1.0+ - how much of the allowance is already eaten.
    function consumedFraction() {
        var allowance = allowanceKcal();
        if (allowance <= 0) {
            return consumedKcal() > 0 ? 1.0 : 0.0;
        }
        return consumedKcal().toFloat() / allowance;
    }

    // ---- sync bookkeeping -----------------------------------------------

    function pendingSync() {
        return _syncedClicks != _clicks;
    }

    function markSynced() {
        _syncedClicks = _clicks;
        save();
    }

    function payload() {
        return {
            "source" => "garmin",
            "date" => _day,
            "sets" => _clicks,
            "kcalPerSet" => kcalPerClick,
            "consumedKcal" => consumedKcal(),
            "activeKcal" => activeCalories(),
            "totalKcal" => totalCalories(),
            "baseBudgetKcal" => dailyBudget,
            "allowanceKcal" => allowanceKcal(),
            "remainingKcal" => remainingKcal(),
            "steps" => steps(),
            "ts" => Time.now().value()
        };
    }
}
