// ================================================================
// GEMINI MODEL RESILIENCE — auto-discovery, failover, graceful degradation
// ================================================================

var GEMINI_MODELS_CACHE_TTL_ = 7 * 24 * 60 * 60 * 1000;

function discoverAvailableModels_(forceRefresh) {
  var props = PropertiesService.getScriptProperties();
  if (!forceRefresh) {
    var cached = props.getProperty("GEMINI_MODELS_CACHE");
    var cacheTs = parseInt(props.getProperty("GEMINI_MODELS_CACHE_TS") || "0", 10);
    if (cached && (Date.now() - cacheTs < GEMINI_MODELS_CACHE_TTL_)) {
      try { return JSON.parse(cached); } catch (e) { /* fall through */ }
    }
  }
  var geminiKey = props.getProperty("GEMINI_API_KEY") || "";
  if (!geminiKey) {
    console.log("discoverAvailableModels_: no GEMINI_API_KEY set");
    return getCachedModelsOrDefaults_(props);
  }
  try {
    var resp = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models?key=" + geminiKey,
      { muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) {
      console.log("discoverAvailableModels_: HTTP " + resp.getResponseCode());
      return getCachedModelsOrDefaults_(props);
    }
    var data = JSON.parse(resp.getContentText());
    var models = (data.models || [])
      .filter(function (m) {
        return m.supportedGenerationMethods &&
               m.supportedGenerationMethods.indexOf("generateContent") !== -1;
      })
      .map(function (m) { return m.name.replace("models/", ""); });
    var ranked = rankModels_(models);
    props.setProperty("GEMINI_MODELS_CACHE", JSON.stringify(ranked));
    props.setProperty("GEMINI_MODELS_CACHE_TS", String(Date.now()));
    console.log("discoverAvailableModels_: found " + ranked.length + " models");
    return ranked;
  } catch (e) {
    console.log("discoverAvailableModels_ error: " + e.toString());
    return getCachedModelsOrDefaults_(props);
  }
}

function rankModels_(models) {
  function tier(name) {
    var n = name.toLowerCase();
    if (n.indexOf("flash-lite") !== -1 || n.indexOf("flash-8b") !== -1) return 1;
    if (n.indexOf("flash") !== -1) return 2;
    if (n.indexOf("gemma") !== -1) return 3;
    if (n.indexOf("pro") !== -1) return 4;
    return 5;
  }
  return models.slice().sort(function (a, b) { return tier(a) - tier(b); });
}

function getCachedModelsOrDefaults_(props) {
  var raw = props.getProperty("GEMINI_MODELS_CACHE");
  if (raw) { try { return JSON.parse(raw); } catch (e) {} }
  return ["gemini-2.5-flash-lite", "gemini-2.5-flash", "gemma-3-27b-it"];
}

function callGeminiWithResilience_(prompt, opts) {
  opts = opts || {};
  var temperature = opts.temperature !== undefined ? opts.temperature : 0.4;
  var maxTokens = opts.maxOutputTokens || 800;
  var geminiKey = opts.geminiKey ||
    PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY") || "";
  if (!geminiKey) { console.log("callGeminiWithResilience_: no API key"); return null; }
  var models = discoverAvailableModels_(false);
  if (!models || models.length === 0) { return null; }
  try {
    var preferred = getAiModel_();
    if (preferred && preferred.model) {
      var idx = models.indexOf(preferred.model);
      if (idx > 0) { models.splice(idx, 1); models.unshift(preferred.model); }
      else if (idx === -1) { models.unshift(preferred.model); }
    }
  } catch (e) {}
  var deadRaw = PropertiesService.getScriptProperties().getProperty("GEMINI_DEAD_MODELS") || "{}";
  var deadMap = {};
  try { deadMap = JSON.parse(deadRaw); } catch (e) { deadMap = {}; }
  var cutoff = Date.now() - 24 * 60 * 60 * 1000;
  models = models.filter(function (m) { return !deadMap[m] || deadMap[m] < cutoff; });
  if (models.length === 0) {
    models = getCachedModelsOrDefaults_(PropertiesService.getScriptProperties());
  }
  var newlyDead = [];
  for (var i = 0; i < models.length && i < 5; i++) {
    var modelName = models[i];
    try {
      var resp = UrlFetchApp.fetch(
        "https://generativelanguage.googleapis.com/v1beta/models/" +
          modelName + ":generateContent?key=" + geminiKey,
        {
          method: "post",
          contentType: "application/json",
          muteHttpExceptions: true,
          payload: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: { temperature: temperature, maxOutputTokens: maxTokens }
          })
        }
      );
      var code = resp.getResponseCode();
      if (code === 404 || code === 400) {
        console.log("callGeminiWithResilience_: " + modelName + " returned " + code);
        newlyDead.push(modelName);
        markModelDead_(modelName);
        continue;
      }
      if (code === 429) { Utilities.sleep(2000); continue; }
      if (code >= 500) { continue; }
      if (code === 200) {
        var json = JSON.parse(resp.getContentText());
        var text = "";
        try { text = json.candidates[0].content.parts[0].text; } catch (e) {}
        if (text) {
          var label = modelToLabel_(modelName);
          recordAiPing_(label);
          return { text: text.trim(), model: modelName, label: label };
        }
      }
    } catch (e) {
      console.log("callGeminiWithResilience_: exception on " + modelName + ": " + e.toString());
    }
  }
  if (newlyDead.length > 0) { try { discoverAvailableModels_(true); } catch (e) {} }
  return null;
}

function modelToLabel_(modelName) {
  var n = modelName.toLowerCase();
  if (n.indexOf("flash-lite") !== -1 || n.indexOf("flash-8b") !== -1) return "gemini-lite";
  if (n.indexOf("flash") !== -1) return "gemini-flash";
  if (n.indexOf("gemma") !== -1) return "gemma";
  return "gemini-flash";
}

function markModelDead_(modelName) {
  try {
    var props = PropertiesService.getScriptProperties();
    var dead = {};
    try { dead = JSON.parse(props.getProperty("GEMINI_DEAD_MODELS") || "{}"); } catch (e) { dead = {}; }
    dead[modelName] = Date.now();
    props.setProperty("GEMINI_DEAD_MODELS", JSON.stringify(dead));
  } catch (e) { console.log("markModelDead_ error: " + e.toString()); }
}

function isAiAvailable_() {
  var key = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY") || "";
  if (!key) return false;
  var models = discoverAvailableModels_(false);
  return models && models.length > 0;
}

function refreshModelRegistry_() {
  var models = discoverAvailableModels_(true);
  PropertiesService.getScriptProperties().deleteProperty("GEMINI_DEAD_MODELS");
  var ui = SpreadsheetApp.getUi();
  if (models && models.length > 0) {
    ui.alert("AI Models Updated",
      "Found " + models.length + " available models:\n\n" +
      models.slice(0, 8).join("\n") + (models.length > 8 ? "\n..." : ""),
      ui.ButtonSet.OK);
  } else {
    ui.alert("AI Models Update",
      "Could not discover models. Check that GEMINI_API_KEY is set in Script Properties.",
      ui.ButtonSet.OK);
  }
}

function generateFallbackEmailBody_(longDate, theme, wotd, speakersList) {
  var body = "Join us for our next Sierra Speakers Toastmasters meeting on " + longDate + "!";
  if (theme) { body += "\n\nOur meeting theme is \"" + theme + "\" \u2014 come ready to be inspired."; }
  if (speakersList) { body += "\n\nPrepared speakers this week: " + speakersList + "."; }
  if (wotd) { body += "\n\nWord of the Day: " + wotd + ". Try to weave it into your speeches and table topics!"; }
  body += "\n\nWe hope to see you there. Guests are always welcome!";
  return body;
}

function showAiUnavailableToast_() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "AI features temporarily unavailable. Using basic templates.",
      "Gemini Status", 8);
  } catch (e) {}
}
