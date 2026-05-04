function setupDailyTriggers() {
    console.log("Check triggers");
    try {
        var allTriggers = ScriptApp.getProjectTriggers();
        for (var i = 0; i < allTriggers.length; i++) {
            var h = safeCall(function () { return allTriggers[i].getHandlerFunction(); });
            if (h === 'codexProjectSummary' || h === 'codexListTriggers') {
                try { ScriptApp.deleteTrigger(allTriggers[i]); } catch (e) { }
            }
        }
        ScriptApp.newTrigger('codexProjectSummary').timeBased().everyDays(1).atHour(16).create();
        ScriptApp.newTrigger('codexProjectSummary').timeBased().everyDays(1).atHour(22).create();
    } catch (err) {
        throw err;
    }
}
function codexListTriggers() {
    var triggers = ScriptApp.getProjectTriggers();
    return {
        ok: true,
        projectId: ScriptApp.getScriptId(),
        projectTimeZone: safeCall(function () { return Session.getScriptTimeZone(); }) || 'Unknown',
        triggerCount: triggers.length,
        triggers: triggers.map(function (trigger) {
            return {
                handlerFunction: safeCall(function () { return trigger.getHandlerFunction(); }),
                eventType: safeEnum(function () { return trigger.getEventType(); }),
                triggerSource: safeEnum(function () { return trigger.getTriggerSource(); }),
                uniqueId: safeString(function () { return trigger.getUniqueId(); })
            };
        })
    };
}
function codexProjectSummary() {
    var triggers = ScriptApp.getProjectTriggers();
    return {
        ok: true,
        projectId: ScriptApp.getScriptId(),
        projectTimeZone: safeCall(function () { return Session.getScriptTimeZone(); }) || 'Unknown',
        triggerCount: triggers.length,
        handlers: uniqueValues_(triggers.map(function (trigger) {
            return safeCall(function () { return trigger.getHandlerFunction(); }) || 'UNKNOWN_HANDLER';
        })),
        sources: uniqueValues_(triggers.map(function (trigger) {
            return safeEnum(function () { return trigger.getTriggerSource(); }) || 'UNKNOWN_SOURCE';
        })),
        eventTypes: uniqueValues_(triggers.map(function (trigger) {
            return safeEnum(function () { return trigger.getEventType(); }) || 'UNKNOWN_EVENT';
        }))
    };
}
function safeCall(fn) {
    try {
        return fn();
    } catch (error) {
        return null;
    }
}
function safeEnum(fn) {
    var value = safeCall(fn);
    return value === null || value === undefined ? null : String(value);
}
function safeString(fn) {
    var value = safeCall(fn);
    return value === null || value === undefined ? null : String(value);
}
function uniqueValues_(items) {
    var seen = {};
    return items.filter(function (item) {
        var key = item || '';
        if (seen[key]) return false;
        seen[key] = true;
        return true;
    });
}