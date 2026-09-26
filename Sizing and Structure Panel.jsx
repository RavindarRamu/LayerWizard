// filename: Sizing and Structure Panel.jsx
//
// Sizing and Structure Panel (classic ExtendScript / ScriptUI)
// Standalone Photoshop script — no UXP, no manifest.json, no HTML/CSS.
// Reads BCOM_VENDOR_SIZINGS.csv (expected in the same folder as this
// script) and shows a floating panel for applying vendor-specified
// canvas size / margins / alignment to the active layer. Applying a row
// first runs the studio's "Layer Stutcure" Photoshop Action, applies the
// CSV sizing, then runs the "Final_Structure" Photoshop Action before
// saving the result as a layered TIFF copy into a "Completed" subfolder
// (the original file itself is left untouched) and closing the document.
//
// Run it from Photoshop: File > Scripts > Browse... and pick this file.
// (To make it always available under File > Scripts, drop a copy into
// Photoshop's own Presets/Scripts folder.)

#target photoshop

//----------------------------------------------------------------------
// CONFIG
//----------------------------------------------------------------------

// Used only to locate this script's own folder (findScriptFolder) — this
// file is guaranteed to always exist, unlike the per-tab CSVs below which
// may not exist yet on a brand-new tab.
var ANCHOR_FILENAME = "BCOM_VENDOR_SIZINGS.csv";
var COMPLETED_FOLDER_NAME = "Completed";

// One tab per sheet, each a fully independent CSV file with the exact same
// 19-column layout (see COL below). To add a new tab: add an entry here
// and, if its CSV file doesn't already exist next to this script, one gets
// created automatically (header row only) the first time the panel runs.
var TABS = [
    { label: "Onfigure", csvFileName: "ONFIGURE.csv" },
    { label: "Vendor Chrome", csvFileName: "BCOM_VENDOR_SIZINGS.csv" },
    { label: "Cheshire", csvFileName: "CHESHIRE.csv" },
    { label: "Still_Off Figure", csvFileName: "STILL_OFF_FIGURE.csv" },
    { label: "Marketplace", csvFileName: "MARKETPLACE.csv" },
    { label: "Photography", csvFileName: "PHOTOGRAPHY.csv" }
];

var CSV_HEADER = "Main_Categoery,Sub_Categorey,Widh,Widh_Unit,Height,Height_Unit,Resolution,Top,Top_Unit,Bottom,Bottom_Unit,Left,Left_Unit,Right,Right Unit,Vertical Allignment,Horizontol Allignment,Selcton Type,COMMENTS";

// Must already be loaded once into Photoshop's own Actions panel (Window >
// Actions > panel menu > Load Actions... > Layer Stutcure.atn) — Photoshop
// remembers a loaded action set across sessions, but there is no reliable
// scripted way to load one on the fly, so this is a one-time manual setup
// step. It builds the full retouching layer structure (Shadow/Reflection
// group, Production group, Smart Object conversions, Select Subject +
// Select and Mask cutout, color correction, and its own trim) before any
// of this script's own CSV-driven sizing runs.
var LAYER_STRUCTURE_ACTION_NAME = "Layer Stutcure";
var LAYER_STRUCTURE_ACTION_SET = "_JAVA Script";

// Same one-time-load requirement as above (Final_Structure.atn, same
// "_JAVA Script" set). Runs after the CSV-driven sizing and right before
// the TIFF save, to apply the studio's final structure/finishing pass to
// the already-sized canvas.
var FINAL_STRUCTURE_ACTION_NAME = "Final_Structure";
var FINAL_STRUCTURE_ACTION_SET = "_JAVA Script";

// User-attribution feature: a required "who is running this" dropdown,
// loaded from a flat list of names next to the script, with the last pick
// cached locally so it's pre-selected next time. As the very last step
// before the TIFF save, every other top-level layer/group gets moved
// inside the "Production" group (built by Final_Structure above) and that
// group is renamed to "<Name>_<MMDD>" for attribution/traceability.
var USERS_CSV_FILENAME = "USERS.csv";
var LAST_USER_CACHE_FILENAME = "last_user.txt";

// Remember-last-active-tab feature: a tiny settings file in Folder.userData
// (NOT next to the script — this is a real per-user OS settings location,
// e.g. "~/Library/Application Support/Adobe/..." on Mac or
// "%APPDATA%/Adobe/..." on Windows, always writable and stable across the
// script/panel's own file location). Format is a single-key JSON object:
//   {"lastActiveTab": "Marketplace"}
// This file is real disk state, not an in-memory variable — it survives
// the panel being closed and reopened AND a full Photoshop restart. See
// loadLastActiveTabSetting()/saveLastActiveTabSetting() below.
var PANEL_SETTINGS_FILENAME = "VendorSizingPanelSettings.json";

// Processing log: one shared CSV file next to the script, appended to
// (never rewritten) once per processed image, success or failure — a
// lightweight audit trail across everyone using this panel. See
// logProcessingEvent() below.
var PROCESSING_LOG_FILENAME = "ProcessingLog.csv";
var PROCESSING_LOG_HEADER = "Timestamp,SystemUser,Tab,Main_Categoery,Sub_Categorey,Document,Result,Details";
var PRODUCTION_GROUP_NAME = "Production";
// Left alone at the top level — everything else gets moved into the
// renamed Production group, but this one stays where it is.
var LAYER_NAME_EXCLUDED_FROM_GROUPING = "Layer 0";

// Column order is fixed based on the known sheet layout (headers in the
// source file contain typos like "Widh" / "Horizotal Allignment", so we
// read by position rather than by header name matching).
var COL = {
    mainCategory: 0,
    itemName: 1,
    width: 2,
    height: 4,
    resolution: 6,
    top: 7,
    bottom: 9,
    left: 11,
    right: 13,
    vAlign: 15,
    hAlign: 16,
    selectionType: 17,
    comments: 18
};

//----------------------------------------------------------------------
// CSV PARSING
//----------------------------------------------------------------------

function splitCsvLine(line) {
    var fields = [];
    var current = "";
    var inQuotes = false;

    for (var i = 0; i < line.length; i++) {
        var ch = line.charAt(i);

        if (ch === '"') {
            inQuotes = !inQuotes;
        } else if (ch === "," && !inQuotes) {
            fields.push(current);
            current = "";
        } else {
            current += ch;
        }
    }
    fields.push(current);
    return fields;
}

function toNumberOrNull(text) {
    if (text === undefined || text === null) return null;
    var trimmed = String(text).replace(/^\s+|\s+$/g, "");
    if (trimmed === "") return null;
    var n = parseFloat(trimmed);
    return isNaN(n) ? null : n;
}

function toTextOrNull(text) {
    if (text === undefined || text === null) return null;
    var trimmed = String(text).replace(/^\s+|\s+$/g, "");
    return trimmed === "" ? null : trimmed;
}

function parseCsv(text) {
    // Strip a UTF-8 BOM if present (common when exported from Excel).
    if (text.charCodeAt(0) === 0xFEFF) {
        text = text.substring(1);
    }

    var lines = text.split(/\r\n|\r|\n/);
    var rows = [];

    // Skip the header row (index 0).
    for (var i = 1; i < lines.length; i++) {
        var line = lines[i];
        if (line.replace(/,/g, "").replace(/^\s+|\s+$/g, "") === "") {
            continue; // blank trailer row
        }

        var fields = splitCsvLine(line);
        var mainCategory = toTextOrNull(fields[COL.mainCategory]);
        if (!mainCategory) {
            continue; // no category on this row — skip
        }

        rows.push({
            mainCategory: mainCategory,
            label: toTextOrNull(fields[COL.itemName]) || "Item",
            width: toNumberOrNull(fields[COL.width]),
            height: toNumberOrNull(fields[COL.height]),
            resolution: toNumberOrNull(fields[COL.resolution]) || 300,
            margins: {
                // Blank margins are treated as 0 (full canvas is the
                // usable area).
                top: toNumberOrNull(fields[COL.top]) || 0,
                bottom: toNumberOrNull(fields[COL.bottom]) || 0,
                left: toNumberOrNull(fields[COL.left]) || 0,
                right: toNumberOrNull(fields[COL.right]) || 0
            },
            vAlign: toTextOrNull(fields[COL.vAlign]) || "Center",
            hAlign: toTextOrNull(fields[COL.hAlign]) || "Center",
            selectionType: toTextOrNull(fields[COL.selectionType]),
            comments: toTextOrNull(fields[COL.comments])
        });
    }

    return rows;
}

// Groups flat rows into Main Category -> [items].
function groupRows(rows) {
    var mainMap = {};
    var mainOrder = [];

    for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        if (!mainMap[row.mainCategory]) {
            mainMap[row.mainCategory] = [];
            mainOrder.push(row.mainCategory);
        }
        mainMap[row.mainCategory].push(row);
    }

    var result = [];
    for (var m = 0; m < mainOrder.length; m++) {
        result.push({ name: mainOrder[m], items: mainMap[mainOrder[m]] });
    }
    return result;
}

// Known deployment locations for this specific studio setup, kept as a
// last-resort fallback below $.fileName — some execution contexts (e.g.
// a script run by injecting its text rather than executing the file
// directly) leave $.fileName blank, which would otherwise silently send
// every lookup straight to the disruptive "locate the file" picker.
// These are Mac paths (this is where the script is deployed today); they
// simply never match on Windows (Folder.exists is false), which is
// harmless — the primary $.fileName-based lookup above already works
// identically on both platforms for normal File > Scripts > Browse
// usage. Add the Windows deployment path(s) here too (e.g.
// "C:/Users/<name>/Documents/.../vendor-sizing-uxp-plugin") once this is
// installed on a Windows machine, to get the same fallback robustness
// there.
var KNOWN_SCRIPT_FOLDERS = [
    "/Users/ravramu/Documents/_JAVA Script/TEST_11_09_2026/vendor-sizing-uxp-plugin",
    "/Users/ravramu/Documents/_JAVA Script/TEST_11_09_2026/_JAVA_Extent_Script"
];

// Resolves this script's own folder — normally via $.fileName, falling
// back to the known deployment locations above if that's unavailable or
// doesn't actually contain the CSV, before finally giving up (callers
// fall back further to a manual file picker as needed).
function findScriptFolder() {
    try {
        var viaDollarFileName = new File($.fileName).parent;
        if (viaDollarFileName && new File(viaDollarFileName.fsName + "/" + ANCHOR_FILENAME).exists) {
            return viaDollarFileName;
        }
    } catch (e) { /* $.fileName unavailable in this execution context */ }

    for (var i = 0; i < KNOWN_SCRIPT_FOLDERS.length; i++) {
        var folder = new Folder(KNOWN_SCRIPT_FOLDERS[i]);
        if (folder.exists && new File(folder.fsName + "/" + ANCHOR_FILENAME).exists) {
            return folder;
        }
    }
    return null;
}

// Loads and parses one tab's CSV file, expected right next to this script.
// A brand-new tab's CSV (added to TABS but never populated yet) won't
// exist on disk yet — rather than popping a disruptive "locate file"
// dialog for that, it's created on the spot with just the header row, so
// the tab just opens empty and ready to be filled in. Only falls back to
// a file picker if the script folder itself can't be found at all.
function loadCsvRows(fileName) {
    var scriptFolder = findScriptFolder();
    var csvFile = scriptFolder ? new File(scriptFolder.fsName + "/" + fileName) : null;

    if (csvFile && !csvFile.exists) {
        csvFile.encoding = "UTF-8";
        csvFile.open("w");
        csvFile.write(CSV_HEADER + "\r\n");
        csvFile.close();
    }

    if (!csvFile || !csvFile.exists) {
        csvFile = File.openDialog("Locate " + fileName, "CSV files:*.csv,All files:*.*");
        if (!csvFile) return []; // user cancelled — treat this tab as empty
    }

    csvFile.encoding = "UTF-8";
    if (!csvFile.open("r")) {
        throw new Error("Couldn't open " + csvFile.fsName);
    }
    var text = csvFile.read();
    csvFile.close();
    return parseCsv(text);
}

// Loads the flat list of user names for the "who is running this" dropdown
// — one name per line, header row skipped. Auto-creates an empty file
// (header only) next to the script if it doesn't exist yet, the same way
// a brand-new tab's CSV does, rather than popping a file picker.
function loadUserNames() {
    var scriptFolder = findScriptFolder();
    var usersFile = scriptFolder ? new File(scriptFolder.fsName + "/" + USERS_CSV_FILENAME) : null;

    if (usersFile && !usersFile.exists) {
        usersFile.encoding = "UTF-8";
        usersFile.open("w");
        usersFile.write("Name\r\n");
        usersFile.close();
    }
    if (!usersFile || !usersFile.exists) return [];

    usersFile.encoding = "UTF-8";
    if (!usersFile.open("r")) return [];
    var text = usersFile.read();
    usersFile.close();

    if (text.charCodeAt(0) === 0xFEFF) text = text.substring(1);
    var lines = text.split(/\r\n|\r|\n/);
    var names = [];
    for (var i = 1; i < lines.length; i++) { // skip header row
        var name = lines[i].replace(/^\s+|\s+$/g, "");
        if (name) names.push(name);
    }
    return names;
}

// Remembers the last-picked user name in a small local text file next to
// the script, so it's pre-selected the next time the panel opens on this
// machine. This is per-machine, not per-Windows/Mac-login — if more than
// one person shares the same computer, it'll suggest whoever ran it last.
function loadLastUsedName() {
    var scriptFolder = findScriptFolder();
    var cacheFile = scriptFolder ? new File(scriptFolder.fsName + "/" + LAST_USER_CACHE_FILENAME) : null;
    if (!cacheFile || !cacheFile.exists) return null;

    cacheFile.encoding = "UTF-8";
    if (!cacheFile.open("r")) return null;
    var text = cacheFile.read();
    cacheFile.close();

    var name = text.replace(/^\s+|\s+$/g, "");
    return name || null;
}

function saveLastUsedName(name) {
    var scriptFolder = findScriptFolder();
    if (!scriptFolder) return;
    var cacheFile = new File(scriptFolder.fsName + "/" + LAST_USER_CACHE_FILENAME);
    cacheFile.encoding = "UTF-8";
    if (cacheFile.open("w")) {
        cacheFile.write(name);
        cacheFile.close();
    }
}

//----------------------------------------------------------------------
// REMEMBER LAST ACTIVE TAB
//----------------------------------------------------------------------
// See PANEL_SETTINGS_FILENAME above for the file location and format.
// Every function here is try/catch-wrapped: a missing, corrupted, or
// unwritable settings file must never crash the panel or block it from
// opening — it just falls back to the default (first) tab and logs the
// issue via $.writeln() instead of alert(), since this is a quiet
// convenience feature, not something the user needs to be interrupted for.

function getPanelSettingsFile() {
    return new File(Folder.userData.fsName + "/" + PANEL_SETTINGS_FILENAME);
}

// Returns the last-active tab's label string, or null if there is none yet
// (first-ever run), the file is missing/unreadable, or its content doesn't
// parse. Deliberately avoids relying on a JSON.parse/stringify polyfill
// (not guaranteed present in every ExtendScript host) — the settings file
// only ever holds this one key, so a small regex is enough and matches the
// simple string-parsing style already used elsewhere in this file (CSV
// parsing, etc.) rather than pulling in a JSON library for one field.
function loadLastActiveTabSetting() {
    try {
        var settingsFile = getPanelSettingsFile();
        if (!settingsFile.exists) return null;

        settingsFile.encoding = "UTF-8";
        if (!settingsFile.open("r")) return null;
        var text = settingsFile.read();
        settingsFile.close();

        if (text.charCodeAt(0) === 0xFEFF) text = text.substring(1);
        var match = text.match(/"lastActiveTab"\s*:\s*"([^"]*)"/);
        return (match && match[1]) ? match[1] : null;
    } catch (e) {
        $.writeln("loadLastActiveTabSetting: couldn't read \"" + PANEL_SETTINGS_FILENAME + "\" — " + errorMessage(e));
        return null;
    }
}

// Persists the given tab label (e.g. "Marketplace") to disk immediately —
// called every time the active tab changes, via attachTabChangeListener
// below, so the very latest tab is always what's on disk even if
// Photoshop or the panel closes without any other save happening first.
function saveLastActiveTabSetting(tabName) {
    try {
        if (!tabName) return;
        var settingsFile = getPanelSettingsFile();
        settingsFile.encoding = "UTF-8";
        if (!settingsFile.open("w")) {
            $.writeln("saveLastActiveTabSetting: couldn't open \"" + PANEL_SETTINGS_FILENAME + "\" for writing.");
            return;
        }
        settingsFile.write("{\"lastActiveTab\": \"" + String(tabName).replace(/\\/g, "\\\\").replace(/"/g, "\\\"") + "\"}");
        settingsFile.close();
    } catch (e) {
        $.writeln("saveLastActiveTabSetting: couldn't write \"" + PANEL_SETTINGS_FILENAME + "\" — " + errorMessage(e));
    }
}

// Called once during buildUI(), BEFORE the panel is shown: reads the
// settings file and selects whichever tab was last active, matching by
// label against `tabList` (the TABS config array — same order/labels the
// tabs were built from). Falls back to the first tab (index 0) — without
// throwing — when there's no stored value yet, the file is missing/
// unreadable, or it names a tab that no longer exists. Returns the
// now-selected tab element (or null if panelRef has no tabs at all) so the
// caller can also set its own activeTabController from it, the same way
// the previous hardcoded "select the first tab" line did.
function restorePanelToLastTab(panelRef, tabList) {
    if (!panelRef || !panelRef.children || panelRef.children.length === 0) return null;

    var lastTabName = loadLastActiveTabSetting();
    var matchedIndex = -1;
    if (lastTabName) {
        for (var i = 0; i < tabList.length; i++) {
            if (tabList[i].label === lastTabName) {
                matchedIndex = i;
                break;
            }
        }
        if (matchedIndex === -1) {
            $.writeln("restorePanelToLastTab: stored tab \"" + lastTabName + "\" no longer exists — defaulting to the first tab.");
        }
    }
    if (matchedIndex === -1 || matchedIndex >= panelRef.children.length) matchedIndex = 0;

    panelRef.selection = panelRef.children[matchedIndex];
    return panelRef.children[matchedIndex];
}

// Wraps whatever onChange handler panelRef already has (preserving it —
// this file's own context-bar/footer-refresh logic keeps running exactly
// as before) and additionally invokes onTabChanged(tabName) after every
// user-driven tab switch, so callers can hook in side effects like
// persisting the new tab without touching the existing handler's code.
function attachTabChangeListener(panelRef, onTabChanged) {
    var previousOnChange = panelRef.onChange;
    panelRef.onChange = function () {
        if (previousOnChange) previousOnChange.call(this);
        var selectedTab = panelRef.selection;
        var tabName = selectedTab ? selectedTab.text : null;
        if (tabName) onTabChanged(tabName);
    };
}

//----------------------------------------------------------------------
// ERRORS
//----------------------------------------------------------------------

function errorMessage(err) {
    if (!err) return "Unknown error";
    if (typeof err === "string") return err;
    if (err.message) return err.message;
    try {
        return String(err);
    } catch (e) {
        return "Unknown error";
    }
}

//----------------------------------------------------------------------
// CORE IMAGE LOGIC — layer structure action -> proportional scale -> canvas expand
//----------------------------------------------------------------------
//
// Full pipeline per row, matching the vendor sizing sheet exactly:
//
//   0. LAYER STRUCTURE — runs the studio's own "Layer Stutcure" Photoshop
//               Action first (see LAYER_STRUCTURE_ACTION_NAME/SET above),
//               which builds the whole retouching structure: unlocks the
//               Background layer, separates Shadow/Reflection elements
//               into their own group, runs Select Subject + Select and
//               Mask to cut the product out, converts every relevant
//               layer to a Smart Object, builds the Production group
//               with its white background and color-correction layers,
//               and trims the document down to that cutout's real
//               extent — shadow included, since it's a real, deliberate
//               part of the structure the action builds rather than an
//               incidental crop this script would have to guess at.
//               This is what actually protects shadow detail through the
//               scale step below: everything of consequence is already a
//               Smart Object before any resampling happens.
//   1. SCALE  — availableWidth/Height = target canvas minus that row's
//               margins. The action's own trimmed result is scaled by
//               ONE uniform factor — min(availableWidth/trimW,
//               availableHeight/trimH) — so it's never stretched or
//               distorted, using whichever dimension hits its limit
//               first. Preserve Details 2.0 is used when upscaling,
//               Bicubic Sharper when downscaling.
//   2. EXPAND — the canvas (never the image) is expanded to the exact
//               target Width x Height, via two whole-document resizeCanvas
//               calls rather than any per-layer move — this repositions
//               every layer AND every extra alpha/spot channel together,
//               automatically and identically, since canvas resize is
//               never layer-scoped. Whichever dimension the scale factor
//               was limited by now exactly fills its available space (so
//               both its margins already match the row exactly); the
//               OTHER dimension is placed by giving it the specified
//               margin on its Alignment side and pushing the leftover
//               slack to the opposite side.
//   2.5 GUIDES — one guide dropped in from each edge of the now-final
//               canvas by that row's own CSV margin (Top/Bottom/Left/
//               Right), same as View > Guides > New Guide Layout — see
//               applyVariantCore's GUIDES step below.
//   3. RESOLUTION — set to the row's Resolution (300 ppi) with no further
//               resampling — this only rewrites the ppi metadata.
//
// Selection Type (CSV column R) documents which isolation method a row
// expects; the Layer Stutcure action's own Select Subject + Select and
// Mask step is used uniformly here regardless of that column's exact
// text, per the brief's "e.g. subject/object-based selection logic".
//
// Onfigure tab only: every row (MODEL_FULL/MODEL_TOP/MODEL_CP_KNEE/
// MODEL_CP_BOTTOM) additionally needs a GUIDE-BASED CROP step BEFORE Layer
// Stutcure — see cropUsingGuideLines() below, ported from the studio's own
// standalone photoshop-horizontal-guide-auto-crop-1200x1500.js. The
// operator places two horizontal guides on the raw image marking the crop
// band; if they haven't, this throws and NOTHING else runs for that row.

// ---------------------------------------------------------------------------
// GUIDE-BASED CROP (Onfigure tab only)
// ---------------------------------------------------------------------------
// Ported from _JAVA_Extent_Script/_Model Crop/photoshop-horizontal-guide-
// auto-crop-1200x1500.js and generalized to any target width/height
// instead of a hardcoded 1200x1500, since Onfigure's own rows span
// 800x1000 / 900x1125 / 1000x1250 / 1200x1500 / 2000x2000 / Detail_Shots
// (1200x1500) — the SAME two guides on the raw image work for whichever
// row the operator picked; only the crop ratio/size changes per row.
function getHorizontalGuidePositions(doc) {
    var positions = [];
    for (var i = 0; i < doc.guides.length; i++) {
        var g = doc.guides[i];
        try {
            if (g.direction === Direction.HORIZONTAL) {
                positions.push(g.coordinate.as("px"));
            }
        } catch (e) { /* an unsupported guide entry — ignore it */ }
    }
    return positions;
}

// Human-readable guide status for the Onfigure tab's UI (see
// refreshGuideStatusLocal in createTabController) — reports on
// app.activeDocument specifically, i.e. what a plain (non-"Apply to all")
// Submit would use. "Apply to all open images" still validates each open
// document's own guides independently at Submit time regardless of what
// this shows, since a batch run can span several documents in different
// states.
function getOnfigureGuideStatusText() {
    if (app.documents.length === 0) {
        return "No document open — open an image, add 2 horizontal guides, then reopen this panel.";
    }
    var activeDoc;
    try {
        activeDoc = app.activeDocument;
    } catch (e) {
        return "No document open — open an image, add 2 horizontal guides, then reopen this panel.";
    }

    var horizontalCount = getHorizontalGuidePositions(activeDoc).length;
    if (horizontalCount >= 2) {
        return "✓ " + horizontalCount + " horizontal guide(s) found on \"" + activeDoc.name + "\" — ready to Submit.";
    }
    return "⚠ No horizontal guides on \"" + activeDoc.name + "\". This panel blocks Photoshop while open — " +
        "close it, add 2 horizontal guides (photoshop-horizontal-guide-auto-crop-1200x1500.js or View > New Guide…), then reopen.";
}

// Requires at least two horizontal guides already on `doc` — throws (rather
// than alert()ing) if they're missing or invalid, the same way every other
// validation failure in this file surfaces through the panel's own status
// bar (see the Submit button's catch block) instead of a blocking native
// dialog, so a batch "Apply to all open images" run can still report which
// specific documents failed instead of stopping on the first alert.
function cropUsingGuideLines(doc, targetWidthPx, targetHeightPx) {
    var guidePositions = getHorizontalGuidePositions(doc);
    if (guidePositions.length < 2) {
        throw new Error(
            "This Onfigure row needs two horizontal guides marking the crop band before it can run. " +
            "Add them (via photoshop-horizontal-guide-auto-crop-1200x1500.js or View > New Guide…), then click Submit again."
        );
    }

    guidePositions.sort(function (a, b) { return a - b; });
    var topGuide = guidePositions[0];
    var bottomGuide = guidePositions[guidePositions.length - 1];

    var docW = doc.width.value;
    var docH = doc.height.value;
    var targetRatio = targetWidthPx / targetHeightPx;

    var bandH = bottomGuide - topGuide;
    if (bandH <= 0) {
        throw new Error("The two horizontal guides on this image are in an invalid position — couldn't compute a crop band.");
    }

    var cropH = Math.min(bandH, docH);
    var cropW = cropH * targetRatio;
    if (cropW > docW) {
        cropW = docW;
        cropH = cropW / targetRatio;
    }

    var left = (docW - cropW) / 2;
    var right = left + cropW;
    var top = topGuide + (bandH - cropH) / 2;
    var bottom = top + cropH;

    if (left < 0) { left = 0; right = cropW; }
    if (right > docW) { right = docW; left = docW - cropW; }
    if (top < 0) { top = 0; bottom = cropH; }
    if (bottom > docH) { bottom = docH; top = docH - cropH; }
    if (top < topGuide) { top = topGuide; bottom = top + cropH; }
    if (bottom > bottomGuide) { bottom = bottomGuide; top = bottom - cropH; }

    doc.crop([
        UnitValue(left, "px"),
        UnitValue(top, "px"),
        UnitValue(right, "px"),
        UnitValue(bottom, "px")
    ]);

    // Clamping above can leave the crop a hair off the exact target ratio
    // (e.g. squeezed against the guide band) — this corrects it exactly
    // like the original standalone script did, just against this row's own
    // target size instead of a hardcoded 1200x1500.
    if (doc.width.value !== targetWidthPx || doc.height.value !== targetHeightPx) {
        doc.resizeImage(
            UnitValue(targetWidthPx, "px"),
            UnitValue(targetHeightPx, "px"),
            null,
            ResampleMethod.BICUBIC
        );
    }
}

function applyVariantCore(doc, variant) {
    if (!variant.width || !variant.height) {
        throw new Error("This row is missing a width/height value in the CSV.");
    }

    // ---- -1. GUIDE-BASED CROP (Onfigure tab only) ----
    // Runs before Layer Stutcure so the action always receives an
    // already-framed image. Throws (stopping everything below, including
    // Layer Stutcure) if the operator hasn't placed the two required
    // horizontal guides yet.
    if (variant.tabLabel === "Onfigure") {
        cropUsingGuideLines(doc, variant.width, variant.height);
    }

    // ---- 0. LAYER STRUCTURE ----
    try {
        app.doAction(LAYER_STRUCTURE_ACTION_NAME, LAYER_STRUCTURE_ACTION_SET);
    } catch (e) {
        throw new Error(
            "Couldn't run the \"" + LAYER_STRUCTURE_ACTION_NAME + "\" action (set \"" + LAYER_STRUCTURE_ACTION_SET + "\"). " +
            "Load it once via Window > Actions > panel menu > Load Actions…, then try again. " +
            "Original error: " + errorMessage(e)
        );
    }
    try { doc.selection.deselect(); } catch (e) { /* no active selection left — fine */ }

    // ---- 0.5 REPOSITION "Rotated Reflection" ----
    // The Layer Stutcure action's own recorded move for this layer assumes
    // whatever image size it was recorded against, so it doesn't land
    // flush against every product's actual bottom edge. This snaps its top
    // edge to the "Image" layer's real bottom edge for THIS image —
    // vertical only, horizontal position is left exactly as the action
    // built it. findLayerByName searches the whole layer tree, so it finds
    // both regardless of which group the action nested them inside.
    var imageLayerForReflection = findLayerByName(doc, "Image");
    if (!imageLayerForReflection) {
        throw new Error("Couldn't find the \"Image\" layer after running \"" + LAYER_STRUCTURE_ACTION_NAME + "\" — can't reposition \"Rotated Reflection\" against it.");
    }
    var rotatedReflectionLayer = findLayerByName(doc, "Rotated Reflection");
    if (!rotatedReflectionLayer) {
        throw new Error("Couldn't find a \"Rotated Reflection\" layer after running \"" + LAYER_STRUCTURE_ACTION_NAME + "\".");
    }
    var reflectionDeltaY = imageLayerForReflection.bounds[3].value - rotatedReflectionLayer.bounds[1].value;
    if (reflectionDeltaY !== 0) {
        rotatedReflectionLayer.translate(0, reflectionDeltaY);
    }

    var layer = doc.activeLayer;
    if (!layer) {
        throw new Error("The layer structure action left no active layer to process.");
    }
    if (layer.typename === "LayerSet") {
        throw new Error("The layer structure action left a group active — expected a single layer.");
    }

    var wasLocked = false;
    try { wasLocked = layer.allLocked; } catch (e) { /* not all layer kinds expose this */ }

    try {
        try { layer.allLocked = false; } catch (e) { /* ignore */ }

        // ---- 1. SCALE ----
        var m = variant.margins;
        var availableWidth = variant.width - (m.left + m.right);
        var availableHeight = variant.height - (m.top + m.bottom);
        if (availableWidth <= 0 || availableHeight <= 0) {
            throw new Error("This row's margins leave no usable area inside the " + variant.width + "x" + variant.height + " canvas.");
        }

        var trimmedW = doc.width.value;
        var trimmedH = doc.height.value;
        var scale = Math.min(availableWidth / trimmedW, availableHeight / trimmedH);

        if (scale !== 1) {
            // Preserve Details 2.0 when enlarging, Bicubic Sharper when
            // reducing — never plain Bicubic — and always a single
            // uniform factor for both dimensions, never independent
            // width/height resizing, so the product is never distorted.
            var resampleMethod = scale > 1 ? ResampleMethod.PRESERVEDETAILS : ResampleMethod.BICUBICSHARPER;
            doc.resizeImage(
                UnitValue(Math.round(trimmedW * scale), "px"),
                UnitValue(Math.round(trimmedH * scale), "px"),
                doc.resolution,
                resampleMethod
            );
        }

        // ---- 2. EXPAND ----
        // The trimmed+scaled product's own edges already coincide with
        // the canvas edges (the layer structure action's own trim). The
        // axis the scale factor was limited by already exactly fills
        // its available space (extra == 0 there, so that axis's near-side
        // gap reduces to exactly its own CSV margin on both sides); on
        // the other axis, "Left"/"Top" alignment keeps the product flush
        // to its own margin and pushes all the leftover slack to the far
        // side, "Right"/"Bottom" does the reverse, and "Center" splits
        // the leftover evenly — so every row's margins land exactly
        // where specified, even when top/bottom or left/right differ.
        var scaledW = doc.width.value;
        var scaledH = doc.height.value;
        var extraW = Math.max(availableWidth - scaledW, 0);
        var extraH = Math.max(availableHeight - scaledH, 0);
        var hAlign = String(variant.hAlign || "Center").toLowerCase();
        var vAlign = String(variant.vAlign || "Center").toLowerCase();

        var leftGap;
        if (hAlign === "right") {
            leftGap = m.left + extraW;
        } else if (hAlign === "center") {
            leftGap = m.left + extraW / 2;
        } else { // "left" (or unrecognized)
            leftGap = m.left;
        }

        var topGap;
        if (vAlign === "bottom") {
            topGap = m.top + extraH;
        } else if (vAlign === "center") {
            topGap = m.top + extraH / 2;
        } else { // "top" (or unrecognized)
            topGap = m.top;
        }

        // Two resizeCanvas calls, never a per-layer/per-channel move: a
        // canvas resize is a whole-document operation, so it repositions
        // every layer AND every extra alpha/spot channel identically and
        // automatically — there is no equivalent single-layer-scoped
        // "translate" step to fall out of sync. First pad only the near
        // (left/top) side by anchoring to the FAR corner (BOTTOMRIGHT),
        // which keeps the content's bottom-right corner fixed while
        // adding exactly leftGap/topGap of new space on the other two
        // sides; then extend out to the final exact target size,
        // anchoring TOPLEFT this time to hold the now-correctly-placed
        // content still while the remaining right/bottom margin is added.
        doc.resizeCanvas(
            UnitValue(scaledW + leftGap, "px"),
            UnitValue(scaledH + topGap, "px"),
            AnchorPosition.BOTTOMRIGHT
        );
        doc.resizeCanvas(
            UnitValue(variant.width, "px"),
            UnitValue(variant.height, "px"),
            AnchorPosition.TOPLEFT
        );

        // ---- 2.5 GUIDES ----
        // Mirrors View > Guides > New Guide Layout: one guide inset from
        // each edge by that row's own CSV margin. Common to all 6 tabs,
        // since every tab's CSV shares the same Top/Bottom/Left/Right
        // margin columns (COL.top/bottom/left/right -> variant.margins).
        // Guides are dropped at the FINAL canvas size set just above, so
        // they land at the same spots the margins describe regardless of
        // this row's Alignment. Cleared first so re-running a row doesn't
        // pile up duplicate guides.
        try { doc.guides.removeAll(); } catch (e) { /* no guides yet — fine */ }
        doc.guides.add(Direction.HORIZONTAL, UnitValue(m.top, "px"));
        doc.guides.add(Direction.HORIZONTAL, UnitValue(variant.height - m.bottom, "px"));
        doc.guides.add(Direction.VERTICAL, UnitValue(m.left, "px"));
        doc.guides.add(Direction.VERTICAL, UnitValue(variant.width - m.right, "px"));

        // ---- 3. RESOLUTION ----
        // Resolution-only change: passing undefined width/height with
        // ResampleMethod.NONE mirrors "Resample: unchecked" in the Image
        // Size dialog — it rewrites only the ppi metadata, no resampling.
        doc.resizeImage(undefined, undefined, variant.resolution, ResampleMethod.NONE);
    } finally {
        try { layer.allLocked = wasLocked; } catch (e) { /* ignore */ }
    }
}

// Saves `doc` as a TIFF, same base filename, into a "Completed" subfolder
// next to the document's own file.
function saveDocumentAsTiffCore(doc) {
    var docFile;
    try {
        docFile = doc.fullName;
    } catch (e) {
        throw new Error("This document has never been saved, so there's no folder to save the TIFF into.");
    }

    var dirFolder = docFile.parent;
    var baseName = docFile.name.replace(/\.[^.]+$/, "");

    var completedFolder = new Folder(dirFolder.fsName + "/" + COMPLETED_FOLDER_NAME);
    if (!completedFolder.exists) {
        completedFolder.create();
    }

    var destFile = new File(completedFolder.fsName + "/" + baseName + ".tif");

    var tiffOptions = new TiffSaveOptions();
    tiffOptions.byteOrder = ByteOrder.IBM;
    tiffOptions.imageCompression = TIFFEncoding.TIFFLZW;
    tiffOptions.layerCompression = LayerCompression.ZIP;
    tiffOptions.layers = true;
    tiffOptions.embedColorProfile = true;

    doc.saveAs(destFile, tiffOptions, true, Extension.LOWERCASE);
}

// Runs `fn` with ruler units forced to pixels and dialogs suppressed,
// then restores both — so bounds/width/height read as plain pixel
// numbers and no TIFF-options dialog can pop up mid-batch.
function withPixelRulerUnits(fn) {
    var previousRulerUnits = app.preferences.rulerUnits;
    var previousDisplayDialogs = app.displayDialogs;
    app.preferences.rulerUnits = Units.PIXELS;
    app.displayDialogs = DialogModes.NO;
    try {
        fn();
    } finally {
        app.preferences.rulerUnits = previousRulerUnits;
        app.displayDialogs = previousDisplayDialogs;
    }
}

// Applies the variant to `doc`, saves a layered TIFF copy to Completed/,
// then closes WITHOUT saving the original file — it's left exactly as
// it was on disk; the TIFF in Completed/ is the deliverable. Both steps
// must happen INSIDE withPixelRulerUnits, not after it: its finally
// block restores app.displayDialogs before returning, and running the
// close afterward — with dialogs back on — risked a real interactive
// "Save changes?" prompt instead of resolving silently.
function applyVariant(doc, variant, userName) {
    withPixelRulerUnits(function () {
        applyVariantCore(doc, variant);

        // Belt-and-braces: playback-specific dialog suppression is a
        // separate switch from app.displayDialogs above — a step inside
        // the action itself (e.g. one recorded with its dialog toggle on)
        // can still pop a native dialog during doAction() even with
        // displayDialogs set to NO, unless this is also set.
        var previousPlaybackDisplayDialogs = app.playbackDisplayDialogs;
        app.playbackDisplayDialogs = DialogModes.NO;
        try {
            app.doAction(FINAL_STRUCTURE_ACTION_NAME, FINAL_STRUCTURE_ACTION_SET);
        } catch (e) {
            throw new Error(
                "Couldn't run the \"" + FINAL_STRUCTURE_ACTION_NAME + "\" action (set \"" + FINAL_STRUCTURE_ACTION_SET + "\"). " +
                "Load it once via Window > Actions > panel menu > Load Actions…, then try again. " +
                "Original error: " + errorMessage(e)
            );
        } finally {
            app.playbackDisplayDialogs = previousPlaybackDisplayDialogs;
        }

        // The action must not contain its own Save/Close steps — if it
        // does, the document is gone before the lines below ever run, so
        // no TIFF gets written. (Final_Structure.atn as originally
        // recorded had a trailing Save + Close; those must be deleted
        // from the action in Window > Actions.) There's no real
        // "isValid" property on Document here (it's undefined, not a
        // boolean, while the doc is open) — a stale reference only
        // reveals itself by throwing when touched, so probe with that.
        var docStillOpen = true;
        try { doc.name; } catch (e) { docStillOpen = false; }
        if (!docStillOpen) {
            throw new Error(
                "The \"" + FINAL_STRUCTURE_ACTION_NAME + "\" action closed the document itself " +
                "(it likely still has its own recorded Save/Close steps at the end). Open Window > " +
                "Actions, expand \"" + FINAL_STRUCTURE_ACTION_SET + "\" > \"" + FINAL_STRUCTURE_ACTION_NAME + "\", " +
                "and delete its last Save and Close steps, then try again."
            );
        }

        finalizeProductionGroup(doc, userName);

        saveDocumentAsTiffCore(doc);
        doc.close(SaveOptions.DONOTSAVECHANGES);
    });
}

// Last step before the TIFF save: moves every other top-level layer/group
// into the "Production" group built by Final_Structure, then renames that
// group to "<userName>_<MMDD>" for attribution/traceability across the
// 100+ people using this script.
//
// Moving a layer/group so it becomes a child INSIDE another group uses
// ElementPlacement.PLACEATEND with the target group as the reference —
// this is the standard, documented technique (there is no simple "INSIDE"
// placement that reliably nests an entire group inside another one).
// Verified live against real documents (each time with a history-state
// checkpoint/rollback so nothing was actually left modified): moving an
// ordinary layer straight into a group works with PLACEATEND, but moving
// another *group* the same way throws "Illegal Argument" — a group can
// only be moved by targeting a layer already inside the destination via
// PLACEBEFORE/PLACEAFTER.
// Searches container.layers recursively (into nested groups), so callers
// don't need to know how deep the action nested a given layer.
function findLayerByName(container, name) {
    for (var i = 0; i < container.layers.length; i++) {
        var item = container.layers[i];
        if (item.name === name) return item;
        if (item.typename === "LayerSet") {
            var found = findLayerByName(item, name);
            if (found) return found;
        }
    }
    return null;
}

function finalizeProductionGroup(doc, userName) {
    var productionGroup;
    try {
        productionGroup = doc.layerSets.getByName(PRODUCTION_GROUP_NAME);
    } catch (e) {
        throw new Error(
            "Couldn't find a \"" + PRODUCTION_GROUP_NAME + "\" group in this document — " +
            "the Layer Stutcure/Final_Structure actions are expected to create it."
        );
    }

    // Collect every top-level item except Production itself and the
    // excluded layer (left alone at the top level) BEFORE moving anything
    // — moving while iterating doc.layers shifts indices.
    var itemsToMove = [];
    for (var i = 0; i < doc.layers.length; i++) {
        var item = doc.layers[i];
        if (item !== productionGroup && item.name !== LAYER_NAME_EXCLUDED_FROM_GROUPING) {
            itemsToMove.push(item);
        }
    }

    // Confirmed live (with a history-state rollback, against a real
    // document): an ordinary layer can move straight inside a group with
    // .move(group, PLACEATEND), but Photoshop throws "Illegal Argument"
    // if the thing being moved is itself a LayerSet (group) — a whole
    // group can only be moved by targeting a layer that's already inside
    // the destination, via PLACEBEFORE/PLACEAFTER, never the destination
    // group itself. So: move ordinary layers in first (each one becomes a
    // valid anchor), then move any other groups relative to one of those.
    var plainLayers = [];
    var groupsToMove = [];
    for (var k = 0; k < itemsToMove.length; k++) {
        if (itemsToMove[k].typename === "LayerSet") {
            groupsToMove.push(itemsToMove[k]);
        } else {
            plainLayers.push(itemsToMove[k]);
        }
    }

    for (var p = 0; p < plainLayers.length; p++) {
        plainLayers[p].move(productionGroup, ElementPlacement.PLACEATEND);
    }

    // Anchor for group moves: whatever's already inside Production —
    // either one of the plainLayers just moved in, or (if there were
    // none) something Final_Structure itself already put there.
    var anchorLayer = productionGroup.layers.length > 0 ? productionGroup.layers[0] : null;
    for (var g = 0; g < groupsToMove.length; g++) {
        var groupItem = groupsToMove[g];

        // "Shadow/Reflection" specifically must land directly beneath the
        // "Image" layer Final_Structure already builds inside Production
        // (confirmed live: PLACEAFTER relative to a layer puts the moved
        // item immediately below it) — everything else just needs to end
        // up inside Production somewhere, so it falls back to the generic
        // anchor below.
        if (groupItem.name === "Shadow/Reflection") {
            var imageLayer = findLayerByName(productionGroup, "Image");
            if (imageLayer) {
                groupItem.move(imageLayer, ElementPlacement.PLACEAFTER);
                continue;
            }
        }

        if (!anchorLayer) {
            throw new Error(
                "Couldn't move \"" + groupItem.name + "\" into \"" + PRODUCTION_GROUP_NAME +
                "\" — the group has no existing layer to anchor the move to."
            );
        }
        groupItem.move(anchorLayer, ElementPlacement.PLACEBEFORE);
    }

    var now = new Date();
    var mm = ("0" + (now.getMonth() + 1)).slice(-2);
    var dd = ("0" + now.getDate()).slice(-2);
    productionGroup.name = userName + "_" + mm + dd;
}

// Returns the OS-level login username (e.g. "ravramu"), NOT the name
// picked from the "User" dropdown — ProcessingLog.csv's SystemUser column
// uses this instead, so the log reflects the actual machine account
// regardless of which attribution name someone selected (that picked name
// is still what drives the Name_Date group rename — this is separate).
// Tries $.getenv first (not guaranteed on every ExtendScript host/version),
// then falls back to parsing Folder.userData's path. Never throws; returns
// null if nothing works.
function getSystemUserName() {
    try {
        if (typeof $.getenv === "function") {
            var envName = $.getenv("USERNAME") || $.getenv("USER") || $.getenv("LOGNAME");
            if (envName) return envName;
        }
    } catch (e) { /* $.getenv unavailable on this ExtendScript host — fall through */ }

    try {
        // Folder.userData looks like ".../Users/<name>/Library/..." (Mac)
        // or "C:\Users\<name>\AppData\..." (Windows) — the OS username is
        // always the path segment right after "Users".
        var match = Folder.userData.fsName.match(/[\\\/][Uu]sers[\\\/]([^\\\/]+)/);
        if (match && match[1]) return match[1];
    } catch (e) { /* ignore */ }

    return null;
}

// Appends one row per processed image (success or failure) to a shared CSV
// log next to the script. Never throws and never blocks the real TIFF-
// save/close pipeline — a logging failure (e.g. the shared file is briefly
// locked by someone else's simultaneous write, on a network/synced folder)
// is swallowed and only reported via $.writeln(), exactly like the other
// quiet-convenience features in this file (remember-last-tab/user).
function logProcessingEvent(fields) {
    try {
        var scriptFolder = findScriptFolder();
        if (!scriptFolder) return;
        var logFile = new File(scriptFolder.fsName + "/" + PROCESSING_LOG_FILENAME);
        logFile.encoding = "UTF-8";
        // Without this, File.writeln() falls back to this ExtendScript
        // host's own default line ending — observed here to be a lone "\r"
        // (classic Mac-style), which most tools (Excel/Numbers included)
        // don't reliably treat as a line break, making every row look like
        // one run-on line. "Windows" (\r\n) is read correctly everywhere.
        logFile.lineFeed = "Windows";

        var needsHeader = !logFile.exists;
        if (!logFile.open("a")) {
            $.writeln("logProcessingEvent: couldn't open \"" + PROCESSING_LOG_FILENAME + "\" for appending.");
            return;
        }
        if (needsHeader) {
            logFile.writeln(PROCESSING_LOG_HEADER);
        }

        var now = new Date();
        var pad = function (n) { return ("0" + n).slice(-2); };
        var timestamp = now.getFullYear() + "-" + pad(now.getMonth() + 1) + "-" + pad(now.getDate()) + " " +
            pad(now.getHours()) + ":" + pad(now.getMinutes()) + ":" + pad(now.getSeconds());

        function csvField(value) {
            var text = (value === undefined || value === null) ? "" : String(value);
            return "\"" + text.replace(/"/g, "\"\"") + "\"";
        }

        logFile.writeln([
            csvField(timestamp),
            csvField(getSystemUserName() || "Unknown"),
            csvField(fields.tabLabel),
            csvField(fields.mainCategory),
            csvField(fields.subCategory),
            csvField(fields.documentName),
            csvField(fields.result),
            csvField(fields.details)
        ].join(","));

        logFile.close();
    } catch (e) {
        $.writeln("logProcessingEvent: couldn't write \"" + PROCESSING_LOG_FILENAME + "\" — " + errorMessage(e));
    }
}

// Applies to every open document rather than just the active one. Each
// document is made active in turn since resizeCanvas/resizeImage always
// target app.activeDocument.
function applyVariantToAllOpenDocuments(variant, userName) {
    var originalActive = app.activeDocument;
    var docs = [];
    for (var i = 0; i < app.documents.length; i++) {
        docs.push(app.documents[i]);
    }

    var succeeded = 0;
    var failures = [];

    for (var j = 0; j < docs.length; j++) {
        var doc = docs[j];
        var docTitleForLog = doc.name; // captured before applyVariant closes doc on success
        try {
            app.activeDocument = doc;
            applyVariant(doc, variant, userName);
            succeeded++;
            logProcessingEvent({
                tabLabel: variant.tabLabel, mainCategory: variant.mainCategory,
                subCategory: variant.label, documentName: docTitleForLog, result: "Success", details: ""
            });
        } catch (error) {
            failures.push(docTitleForLog + ": " + errorMessage(error));
            logProcessingEvent({
                tabLabel: variant.tabLabel, mainCategory: variant.mainCategory,
                subCategory: variant.label, documentName: docTitleForLog, result: "Failed", details: errorMessage(error)
            });
        }
    }

    try {
        app.activeDocument = originalActive; // may have been one of the docs just closed — ignore if so
    } catch (e) { /* ignore */ }

    return { succeeded: succeeded, total: docs.length, failures: failures };
}

//----------------------------------------------------------------------
// UI (ScriptUI)
//----------------------------------------------------------------------

var win, statusText;
var ctxCategoryText, ctxItemText;
var applyAllToggle, cancelButton, submitButton, closeButton;
var tabbedPanel;
var activeTabController = null;
var userDropdown;

function getSelectedUserName() {
    return userDropdown && userDropdown.selection ? userDropdown.selection.text : null;
}

function matchesQuery(text, query) {
    return !!text && text.toLowerCase().indexOf(query.toLowerCase()) !== -1;
}

function setStatus(msg, isError) {
    statusText.text = msg;
    try {
        var g = statusText.graphics;
        g.foregroundColor = isError
            ? g.newPen(g.PenType.SOLID_COLOR, [0.89, 0.28, 0.31], 1)
            : g.newPen(g.PenType.SOLID_COLOR, [0.65, 0.65, 0.65], 1);
    } catch (e) { /* older ScriptUI hosts may not support this */ }
}

// Reflects whichever tab is currently active — each tab keeps its own
// selection state entirely independently (see createTabController below),
// so switching tabs just re-reads that tab's own state through
// activeTabController rather than sharing one global selectedCategory /
// selectedVariant across all 6 tabs.
function updateContextBar() {
    ctxCategoryText.text = (activeTabController && activeTabController.getSelectedCategoryName()) || "—";
    ctxItemText.text = (activeTabController && activeTabController.getSelectedVariantLabel()) || "—";
}

function updateFooterButtons() {
    var hasSelection = !!(activeTabController && activeTabController.getSelectedVariant());
    submitButton.enabled = hasSelection && !!getSelectedUserName();
    cancelButton.enabled = hasSelection;
}

// Called by a tab whenever ITS OWN selection changes — only refreshes the
// shared context bar/footer buttons when that tab is the one currently
// showing, so a background tab's state changes never leak into view.
function refreshSharedChromeIfActive(controller) {
    if (controller === activeTabController) {
        updateContextBar();
        updateFooterButtons();
    }
}

// Bundled reference photos live in images/<MAIN_CATEGORY>__<ITEM_LABEL>.png,
// next to this script (all non-alphanumeric characters squashed to a single
// underscore — same convention as the original UXP plugin). Only a subset
// of items have one; everything else falls back to the generated schematic.
// This Photoshop's ScriptUI "image" control never rescales a bitmap — it
// always shows the file at native pixel size, cropped to whatever bounds
// it's given. So the panel doesn't display the original full-resolution
// bundled photos directly; it displays pre-generated thumbnails from this
// subfolder instead, each already letterboxed (full image, white-padded,
// never cropped) to its own CSV row's exact crop ratio at a fixed
// PREVIEW_TARGET_WIDTH — see generate_previews.py in this same folder,
// which must be re-run after adding/replacing a photo in images/. Shared
// across every tab: sanitized keys only collide if two tabs use the exact
// same Main_Categoery + Sub_Categorey, which is harmless (same photo).
var PREVIEWS_FOLDER_NAME = "images/previews";
var PREVIEW_TARGET_WIDTH = 220;

function sanitizeForFilename(text) {
    return String(text).replace(/^\s+|\s+$/g, "").toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

function imageKeyFor(variant) {
    return sanitizeForFilename(variant.mainCategory) + "__" + sanitizeForFilename(variant.label);
}

var previewImagePathCache = {}; // imageKey -> fsName string, or null if confirmed absent

function loadPreviewImagePath(variant) {
    var key = imageKeyFor(variant);
    if (Object.prototype.hasOwnProperty.call(previewImagePathCache, key)) {
        return previewImagePathCache[key];
    }
    var scriptFolder = findScriptFolder();
    var imageFile = scriptFolder ? new File(scriptFolder.fsName + "/" + PREVIEWS_FOLDER_NAME + "/" + key + ".png") : null;
    previewImagePathCache[key] = (imageFile && imageFile.exists) ? imageFile.fsName : null;
    return previewImagePathCache[key];
}

// Fits a box of itemW:itemH aspect ratio (the vendor's own target crop
// ratio — 1:1, 4:5, whatever the CSV row says) inside a containerW x
// containerH area, centered, preserving the full item — never cropped.
// Shared by the schematic draw and the real-photo image control so both
// always show the correct, un-distorted crop ratio for the selected row.
function computeFitBox(containerW, containerH, itemW, itemH) {
    var pad = 10;
    var availW = Math.max(containerW - pad * 2, 1);
    var availH = Math.max(containerH - pad * 2, 1);
    var scale = Math.min(availW / itemW, availH / itemH);
    var boxW = itemW * scale;
    var boxH = itemH * scale;
    return {
        x: (containerW - boxW) / 2,
        y: (containerH - boxH) / 2,
        width: boxW,
        height: boxH
    };
}

// Builds one tab's entire self-contained UI (its own Main_Categoery /
// Sub_Categorey / Preview columns, search field, and detail box) inside
// `tabContainer`, loads that tab's own CSV file, and returns a small
// controller object the shared chrome (context bar, footer Submit/Cancel)
// uses to read or clear whichever variant is selected in this tab. Each
// tab keeps fully independent state — deliberately not sharing one global
// selectedCategory/selectedVariant the way the original one-sheet panel
// did, since two tabs could otherwise clobber each other's selection.
function createTabController(tabContainer, tabConfig) {
    var controller = {}; // populated at the bottom; closures below reference this same object

    var categories = [];
    var selectedCategory = null;
    var selectedVariant = null;
    var mainList, itemsList, searchField;
    var previewToggle, previewCanvas, previewImageControl, previewArea;
    var detailCanvasValue, detailMarginsValue, detailAlignValue, detailResolutionValue, detailCommentsValue;
    var guideStatusText = null; // Onfigure tab only — see refreshGuideStatusLocal below

    tabContainer.orientation = "column";
    tabContainer.alignChildren = ["fill", "top"];
    tabContainer.margins = 10;
    tabContainer.spacing = 8;

    searchField = tabContainer.add("edittext", undefined, "");
    searchField.helpTip = "Search categories or items...";

    // Three columns, side by side and aligned to the same height, named
    // after the CSV's own headers: Main_Categoery / Sub_Categorey / Preview.
    var bodyGroup = tabContainer.add("group");
    bodyGroup.orientation = "row";
    bodyGroup.alignChildren = ["fill", "fill"];
    bodyGroup.spacing = 12;

    var mainCategoryColumn = bodyGroup.add("panel", undefined, "Main_Categoery");
    mainCategoryColumn.alignChildren = ["fill", "fill"];
    mainCategoryColumn.alignment = ["fill", "fill"];
    mainList = mainCategoryColumn.add("listbox", undefined, []);
    mainList.preferredSize = [170, 170];

    var itemsColumn = bodyGroup.add("panel", undefined, "Sub_Categorey");
    itemsColumn.alignChildren = ["fill", "fill"];
    itemsColumn.alignment = ["fill", "fill"];
    itemsList = itemsColumn.add("listbox", undefined, []);
    itemsList.preferredSize = [200, 170];

    var rightPane = bodyGroup.add("panel", undefined, "Preview");
    rightPane.orientation = "column";
    rightPane.alignChildren = ["fill", "top"];
    rightPane.alignment = ["fill", "fill"];
    rightPane.preferredSize.width = 340;

    // Onfigure-only: every row here needs the guide-based crop (see
    // cropUsingGuideLines in applyVariantCore), which requires 2 horizontal
    // guides already on the document BEFORE Submit is clicked. Since this
    // panel is a modal "dialog" — Photoshop itself is blocked while it's
    // open (see buildUI's own comment on why "dialog" was chosen over
    // "palette") — the user can't add guides after opening the panel, so
    // this surfaces the guide status up front (tab activation / launch),
    // not reactively while they work, via refreshGuideStatusLocal below.
    if (tabConfig.label === "Onfigure") {
        guideStatusText = rightPane.add("statictext", undefined, "", { multiline: true });
        guideStatusText.preferredSize = [320, 34];
    }

    previewToggle = rightPane.add("checkbox", undefined, "Preview");
    previewToggle.value = true;

    previewArea = rightPane.add("group");
    // Without this, a plain group defaults to "row" orientation — flowing
    // the schematic canvas and the bundled-photo image control side by
    // side instead of overlapping both at the same spot. With "row" in
    // effect, hiding one of them (switching categories/items, or a window
    // resize) can leave the other's last-painted pixels showing through
    // instead of properly clearing, since the layout treats them as two
    // flow slots rather than one stacked spot. This is a one-time
    // structural setting with no ongoing performance cost — not to be
    // confused with the (reverted) heavier per-resize-tick workaround.
    previewArea.orientation = "stack";
    previewArea.preferredSize = [240, 285];
    previewArea.margins = 0;

    // A real bundled photo is shown via a native ScriptUI "image" control
    // (the graphics object in this Photoshop has no drawImage/newImage
    // support), and the hand-drawn schematic below it otherwise — only one
    // of the two is ever visible at a time (see refreshPreviewLocal()).
    // This Photoshop's "image" control only scales the bitmap to fit at
    // construction time — resizing an existing one via .bounds afterward
    // just clips it at native pixel resolution instead of rescaling — so
    // previewImageControl is deliberately left uncreated here; refreshPreviewLocal()
    // builds a fresh one (with the right fitted bounds) each time a photo
    // needs to be shown.
    previewImageControl = null;

    previewCanvas = previewArea.add("group");
    previewCanvas.alignment = ["fill", "fill"];

    var detailBox = rightPane.add("panel", undefined, "Selected Item");
    detailBox.orientation = "column";
    detailBox.alignChildren = ["fill", "top"];
    detailBox.spacing = 4;

    function addDetailRow(label, multiline) {
        var row = detailBox.add("group");
        row.orientation = "row";
        var labelText = row.add("statictext", undefined, label);
        labelText.preferredSize.width = 95;
        var valueText = row.add("statictext", undefined, "—", multiline ? { multiline: true } : undefined);
        valueText.alignment = ["fill", multiline ? "fill" : "top"];
        if (multiline) valueText.preferredSize = [220, 32];
        return valueText;
    }
    detailCanvasValue = addDetailRow("Canvas");
    detailMarginsValue = addDetailRow("Margins T/B/L/R");
    detailAlignValue = addDetailRow("Align V/H");
    detailResolutionValue = addDetailRow("Resolution");
    // Surfaces the CSV row's own COMMENTS column (column S) — shows
    // whatever note the sheet has for the currently selected Sub_Categorey
    // item, updating every time a different one is clicked.
    detailCommentsValue = addDetailRow("Comments", true);

    function updateDetailBoxLocal(variant) {
        if (!variant) {
            detailCanvasValue.text = "—";
            detailMarginsValue.text = "—";
            detailAlignValue.text = "—";
            detailResolutionValue.text = "—";
            detailCommentsValue.text = "—";
            return;
        }
        var m = variant.margins;
        detailCanvasValue.text = (variant.width && variant.height)
            ? (variant.width + " x " + variant.height + " px")
            : "no size";
        detailMarginsValue.text = m.top + " / " + m.bottom + " / " + m.left + " / " + m.right;
        detailAlignValue.text = variant.vAlign + " / " + variant.hAlign;
        detailResolutionValue.text = variant.resolution + " ppi";
        detailCommentsValue.text = variant.comments || "No comments";
    }

    // Draws the bundled reference photo (if this item has one) or, failing
    // that, a small schematic: the canvas outline, a box for the usable
    // area inside the CSV margins, and a dot marking where hAlign/vAlign
    // would place the product — generated straight from the row's own
    // numbers.
    function drawPreviewCoreLocal() {
        var g = previewCanvas.graphics;
        var w = previewCanvas.size.width;
        var h = previewCanvas.size.height;

        g.newPath();
        g.rectPath(0, 0, w, h);
        g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, [0.09, 0.09, 0.1]));

        var variant = selectedVariant;
        if (!variant || !variant.width || !variant.height) {
            return;
        }

        var box = computeFitBox(w, h, variant.width, variant.height);
        var offsetX = box.x, offsetY = box.y, canvasW = box.width, canvasH = box.height;
        var scale = canvasW / variant.width;

        // Bundled photos are NOT drawn here — this Photoshop's ScriptUI graphics
        // object has no newImage()/drawImage() (confirmed: it throws "not a
        // function"), so a real bitmap can only be shown via a native ScriptUI
        // "image" control, which previewImageControl (built alongside this
        // canvas) handles; refreshPreviewLocal() picks one or the other to display.

        g.newPath();
        g.rectPath(offsetX, offsetY, canvasW, canvasH);
        g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.55, 0.55, 0.58], 1));

        var m = variant.margins;
        var mLeft = m.left * scale;
        var mRight = m.right * scale;
        var mTop = m.top * scale;
        var mBottom = m.bottom * scale;
        var innerW = Math.max(canvasW - mLeft - mRight, 0);
        var innerH = Math.max(canvasH - mTop - mBottom, 0);

        if (innerW > 0 && innerH > 0) {
            g.newPath();
            g.rectPath(offsetX + mLeft, offsetY + mTop, innerW, innerH);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.15, 0.5, 0.92], 1.5));

            var anchorX;
            if (variant.hAlign === "Left") {
                anchorX = offsetX + mLeft + innerW * 0.15;
            } else if (variant.hAlign === "Right") {
                anchorX = offsetX + mLeft + innerW * 0.85;
            } else {
                anchorX = offsetX + mLeft + innerW / 2;
            }

            var anchorY;
            if (variant.vAlign === "Top") {
                anchorY = offsetY + mTop + innerH * 0.15;
            } else if (variant.vAlign === "Bottom") {
                anchorY = offsetY + mTop + innerH * 0.85;
            } else {
                anchorY = offsetY + mTop + innerH / 2;
            }

            var r = 4;
            g.newPath();
            g.ellipsePath(anchorX - r, anchorY - r, r * 2, r * 2);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, [0.95, 0.68, 0.15]));
        }
    }

    previewCanvas.onDraw = function () {
        try {
            drawPreviewCoreLocal();
        } catch (e) {
            // A drawing error here must never propagate — it happens inside
            // ScriptUI's own redraw pass, and letting it escape can silently
            // abort window creation/display instead of just skipping a frame.
        }
    };

    function refreshPreviewLocal() {
        if (!previewToggle.value) return;

        var variant = selectedVariant;
        var imagePath = variant ? loadPreviewImagePath(variant) : null;
        if (imagePath && variant && variant.width && variant.height) {
            try {
                // This Photoshop's ScriptUI "image" control never rescales a
                // bitmap — whatever bounds you give it, it shows the file at
                // its own native pixel size, cropped to fit. So rather than
                // computing an arbitrary display box, this uses the PREVIEW
                // FILE's own known dimensions (generate_previews.py always
                // makes it PREVIEW_TARGET_WIDTH wide, scaled to the item's
                // exact CSV crop ratio) and just centers a same-size box —
                // no scaling is asked of ScriptUI, so none is lost.
                var previewW = PREVIEW_TARGET_WIDTH;
                var previewH = Math.round(PREVIEW_TARGET_WIDTH * variant.height / variant.width);
                var containerW = previewCanvas.size.width;
                var containerH = previewCanvas.size.height;
                var offsetX = Math.max((containerW - previewW) / 2, 0);
                var offsetY = Math.max((containerH - previewH) / 2, 0);
                var bounds = [offsetX, offsetY, offsetX + previewW, offsetY + previewH];

                // Rebuilt from scratch every time: this Photoshop's ScriptUI
                // only applies bounds at the moment an "image" control is
                // constructed — changing .bounds on an existing one afterward
                // just re-crops the native-resolution bitmap in place.
                if (previewImageControl) {
                    previewArea.remove(previewImageControl);
                }
                previewImageControl = previewArea.add("image", bounds, new File(imagePath));
                previewImageControl.visible = true;
                previewCanvas.visible = false;
                return;
            } catch (e) {
                // Corrupt/unreadable image file — fall back to the schematic.
            }
        }

        if (previewImageControl) previewImageControl.visible = false;
        previewCanvas.visible = true;
        try { previewCanvas.notify("onDraw"); } catch (e) { /* ignore */ }
    }

    // No-op on every tab except Onfigure (guideStatusText stays null there).
    function refreshGuideStatusLocal() {
        if (!guideStatusText) return;
        var text = getOnfigureGuideStatusText();
        guideStatusText.text = text;
        try {
            var g = guideStatusText.graphics;
            var isReady = text.charAt(0) === "✓";
            g.foregroundColor = isReady
                ? g.newPen(g.PenType.SOLID_COLOR, [0.3, 0.75, 0.35], 1)
                : g.newPen(g.PenType.SOLID_COLOR, [0.95, 0.68, 0.15], 1);
        } catch (e) { /* older ScriptUI hosts may not support this */ }
    }

    function selectVariantLocal(variant) {
        selectedVariant = variant;
        updateDetailBoxLocal(variant);
        refreshPreviewLocal();
        refreshGuideStatusLocal();
        refreshSharedChromeIfActive(controller);
        setStatus("Selected \"" + variant.label + "\" — click Submit to apply.", false);
    }

    function clearSelectionLocal(statusMsg) {
        selectedVariant = null;
        if (itemsList) itemsList.selection = null;
        updateDetailBoxLocal(null);
        refreshPreviewLocal();
        refreshGuideStatusLocal();
        refreshSharedChromeIfActive(controller);
        if (statusMsg) setStatus(statusMsg, false);
    }

    function renderItemsListLocal(query) {
        clearSelectionLocal(null);
        itemsList.removeAll();
        if (!selectedCategory) return;

        var shown = 0;
        for (var i = 0; i < selectedCategory.items.length; i++) {
            var v = selectedCategory.items[i];
            if (query && !matchesQuery(v.label, query)) continue;

            var dims = (v.width && v.height) ? (v.width + " x " + v.height + " px") : "no size";
            var li = itemsList.add("item", v.label + "   —   " + dims);
            li.variantRef = v;
            if (v.comments) li.helpTip = v.comments;
            shown++;
        }

        if (shown === 0) {
            itemsList.add("item", "No matching variants.");
        }
    }

    function renderMainListLocal(query) {
        var previousCategoryName = selectedCategory ? selectedCategory.name : null;
        mainList.removeAll();

        var visibleCategories = [];
        for (var i = 0; i < categories.length; i++) {
            var cat = categories[i];
            var categoryMatches = !query || matchesQuery(cat.name, query);
            var itemMatches = false;

            if (query && !categoryMatches) {
                for (var j = 0; j < cat.items.length; j++) {
                    if (matchesQuery(cat.items[j].label, query)) { itemMatches = true; break; }
                }
            }

            if (!query || categoryMatches || itemMatches) {
                visibleCategories.push(cat);
            }
        }

        var selectIndex = -1;
        for (var k = 0; k < visibleCategories.length; k++) {
            var entry = visibleCategories[k];
            var li = mainList.add("item", entry.name);
            li.categoryRef = entry;
            if (entry.name === previousCategoryName) selectIndex = k;
        }
        if (selectIndex === -1 && visibleCategories.length > 0) selectIndex = 0;

        if (selectIndex !== -1) {
            mainList.selection = selectIndex;
            selectedCategory = visibleCategories[selectIndex];
        } else {
            selectedCategory = null;
        }
    }

    mainList.onChange = function () {
        var sel = mainList.selection;
        if (sel && sel.categoryRef) {
            selectedCategory = sel.categoryRef;
            renderItemsListLocal(searchField.text);
            refreshSharedChromeIfActive(controller);
        }
    };

    itemsList.onChange = function () {
        var sel = itemsList.selection;
        if (sel && sel.variantRef) {
            selectVariantLocal(sel.variantRef);
        } else {
            clearSelectionLocal(null);
        }
    };

    searchField.onChanging = function () {
        var q = searchField.text;
        renderMainListLocal(q);
        renderItemsListLocal(q);
        refreshSharedChromeIfActive(controller);
    };

    previewToggle.onClick = function () {
        if (!previewToggle.value) {
            if (previewImageControl) previewImageControl.visible = false;
            previewCanvas.visible = false;
        }
        win.layout.layout(true);
        refreshPreviewLocal();
    };

    // ---- load this tab's own CSV and populate its lists ----
    var rows = [];
    try {
        rows = loadCsvRows(tabConfig.csvFileName);
    } catch (e) {
        alert("Couldn't load \"" + tabConfig.csvFileName + "\" for the \"" + tabConfig.label + "\" tab:\n" + errorMessage(e));
    }
    // Tags each row with which tab it came from — applyVariantCore uses
    // this to know an Onfigure row needs the guide-based crop (see
    // cropUsingGuideLines) before anything else runs.
    for (var rowIndex = 0; rowIndex < rows.length; rowIndex++) {
        rows[rowIndex].tabLabel = tabConfig.label;
    }
    categories = groupRows(rows);
    renderMainListLocal("");
    renderItemsListLocal("");
    refreshGuideStatusLocal();

    controller.label = tabConfig.label;
    controller.getSelectedVariant = function () { return selectedVariant; };
    controller.getSelectedCategoryName = function () { return selectedCategory ? selectedCategory.name : null; };
    controller.getSelectedVariantLabel = function () { return selectedVariant ? selectedVariant.label : null; };
    controller.clearSelection = function (msg) { clearSelectionLocal(msg); };
    controller.refreshPreview = function () { refreshPreviewLocal(); };
    controller.refreshGuideStatus = function () { refreshGuideStatusLocal(); };

    return controller;
}

function buildUI() {
    // A floating "palette" window would be the closest ScriptUI equivalent
    // to a real Photoshop panel, but it silently fails to render at all on
    // some Photoshop/OS combinations (no error — the window object is
    // created, it just never appears on screen). "dialog" reliably shows
    // every time, at the cost of being modal: Photoshop itself is blocked
    // while this window is open, so pick the layer/document you want to
    // apply a template to *before* running the script.
    win = new Window("dialog", "Sizing and Structure Panel", undefined, { resizeable: true });
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.margins = 10;
    win.spacing = 6;
    win.preferredSize.width = 790;

    var contextBar = win.add("group");
    contextBar.add("statictext", undefined, "Main_Categoery:");
    ctxCategoryText = contextBar.add("statictext", undefined, "—");
    ctxCategoryText.preferredSize.width = 160;
    contextBar.add("statictext", undefined, "|");
    contextBar.add("statictext", undefined, "Sub_Categorey:");
    ctxItemText = contextBar.add("statictext", undefined, "—");
    ctxItemText.preferredSize.width = 220;

    var instructionText = win.add("statictext", undefined,
        "Pick a category and item, check the preview, then Submit to apply it (runs Layer " +
        "Stutcure → CSV sizing → Final_Structure → TIFF saved to Completed, original left untouched).",
        { multiline: true });
    instructionText.preferredSize = [630, 30];

    // Required "who is running this" attribution dropdown — Submit stays
    // disabled until a name is picked. Pre-selects whatever was picked
    // last on this machine (see LAST_USER_CACHE_FILENAME).
    var userGroup = win.add("group");
    userGroup.orientation = "row";
    userGroup.alignChildren = ["left", "center"];
    userGroup.add("statictext", undefined, "User (required):");
    userDropdown = userGroup.add("dropdownlist", undefined, loadUserNames());
    userDropdown.preferredSize.width = 220;
    if (userDropdown.items.length === 0) {
        userGroup.add("statictext", undefined, "— add names to " + USERS_CSV_FILENAME + " next to this script —");
    }

    var lastUsedName = loadLastUsedName();
    if (lastUsedName) {
        for (var ui = 0; ui < userDropdown.items.length; ui++) {
            if (userDropdown.items[ui].text === lastUsedName) {
                userDropdown.selection = ui;
                break;
            }
        }
    }

    userDropdown.onChange = function () {
        if (userDropdown.selection) saveLastUsedName(userDropdown.selection.text);
        updateFooterButtons();
    };

    // One tabbedpanel, one tab per TABS entry — each tab builds and owns
    // its own Main_Categoery/Sub_Categorey/Preview UI and CSV data via
    // createTabController(); the context bar and Submit/Cancel below are
    // shared chrome that always reflect whichever tab is active.
    tabbedPanel = win.add("tabbedpanel", undefined);
    tabbedPanel.alignment = ["fill", "fill"];
    tabbedPanel.preferredSize = [770, 380];

    for (var i = 0; i < TABS.length; i++) {
        var tabConfig = TABS[i];
        var tab = tabbedPanel.add("tab", undefined, tabConfig.label);
        tab.controller = createTabController(tab, tabConfig);
    }

    // Restores whichever tab was active last time (see "REMEMBER LAST
    // ACTIVE TAB" above) — e.g. if the panel was last left on
    // "Marketplace", it reopens directly on "Marketplace" instead of
    // always resetting to the first tab. Falls back to the first tab on a
    // first-ever run or an unreadable/unrecognized settings file.
    var restoredTab = restorePanelToLastTab(tabbedPanel, TABS);
    activeTabController = restoredTab ? restoredTab.controller : null;
    // Covers the case where the panel launches directly on the Onfigure
    // tab (e.g. it was the last-active tab) — shows guide status right
    // away rather than waiting for the user to switch tabs or pick a row.
    if (activeTabController && activeTabController.refreshGuideStatus) activeTabController.refreshGuideStatus();

    tabbedPanel.onChange = function () {
        var tab = tabbedPanel.selection;
        activeTabController = tab ? tab.controller : null;
        updateContextBar();
        updateFooterButtons();
        setStatus("Ready.", false);
        if (activeTabController && activeTabController.refreshGuideStatus) activeTabController.refreshGuideStatus();
    };
    // Persists the newly active tab to disk every time the user switches
    // tabs, without touching the onChange handler set just above.
    attachTabChangeListener(tabbedPanel, function (tabName) {
        saveLastActiveTabSetting(tabName);
    });

    var footer = win.add("group");
    footer.orientation = "column";
    footer.alignChildren = ["fill", "top"];
    footer.spacing = 8;

    applyAllToggle = footer.add("checkbox", undefined, "Apply this action to all open images");

    var footerButtons = footer.add("group");
    footerButtons.orientation = "row";
    footerButtons.alignment = ["fill", "top"];
    closeButton = footerButtons.add("button", undefined, "Close");
    closeButton.alignment = ["left", "top"];

    var rightButtons = footerButtons.add("group");
    rightButtons.orientation = "row";
    rightButtons.alignment = ["right", "top"];
    cancelButton = rightButtons.add("button", undefined, "Cancel");
    submitButton = rightButtons.add("button", undefined, "Submit");
    cancelButton.enabled = false;
    submitButton.enabled = false;

    statusText = win.add("statictext", undefined, "Loading...");
    statusText.preferredSize = [630, 22];
    statusText.properties = { multiline: true };

    var creditText = win.add("statictext", undefined, "Script Developed by Ravindar Ramu");
    creditText.alignment = ["right", "top"];
    try {
        var creditGraphics = creditText.graphics;
        creditGraphics.foregroundColor = creditGraphics.newPen(creditGraphics.PenType.SOLID_COLOR, [0.55, 0.55, 0.55], 1);
    } catch (e) { /* older ScriptUI hosts may not support this */ }

    // ---- events ----

    closeButton.onClick = function () {
        win.close();
    };

    cancelButton.onClick = function () {
        if (activeTabController) activeTabController.clearSelection("Selection cleared.");
    };

    submitButton.onClick = function () {
        var controller = activeTabController;
        var variant = controller ? controller.getSelectedVariant() : null;
        var userName = getSelectedUserName();
        if (!variant || !userName) return;
        submitButton.enabled = false;
        cancelButton.enabled = false;
        // Only used on the single-document path below — the "Apply to all
        // open images" path logs per-document itself, inside
        // applyVariantToAllOpenDocuments, since that's the only place that
        // knows each individual document's own name/outcome.
        var singleDocNameForLog = null;
        try {
            if (applyAllToggle.value) {
                if (app.documents.length === 0) {
                    throw new Error("No open images to apply to.");
                }
                var result = applyVariantToAllOpenDocuments(variant, userName);
                if (result.failures.length === 0) {
                    setStatus(
                        "Applied \"" + variant.label + "\" to all " + result.total +
                        " open image(s) — each saved, a TIFF copy written to Completed, and closed.",
                        false
                    );
                } else {
                    setStatus(
                        "Applied/saved/closed " + result.succeeded + "/" + result.total + " open image(s) — " +
                        result.failures[0] + (result.failures.length > 1 ? " (+" + (result.failures.length - 1) + " more)" : ""),
                        true
                    );
                }
            } else {
                if (app.documents.length === 0) {
                    throw new Error("Open an image before applying a template.");
                }
                singleDocNameForLog = app.activeDocument.name;
                applyVariant(app.activeDocument, variant, userName);
                logProcessingEvent({
                    tabLabel: variant.tabLabel, mainCategory: variant.mainCategory,
                    subCategory: variant.label, documentName: singleDocNameForLog, result: "Success", details: ""
                });
                setStatus(
                    "Applied \"" + variant.label + "\" (" +
                    variant.width + " x " + variant.height + " px @ " + variant.resolution + " ppi)" +
                    " — saved, TIFF copy written to Completed, and closed.",
                    false
                );
            }
        } catch (error) {
            if (!applyAllToggle.value && singleDocNameForLog) {
                logProcessingEvent({
                    tabLabel: variant.tabLabel, mainCategory: variant.mainCategory,
                    subCategory: variant.label, documentName: singleDocNameForLog, result: "Failed", details: errorMessage(error)
                });
            }
            setStatus("Failed: " + variant.label + " — " + errorMessage(error), true);
        } finally {
            updateFooterButtons();
        }
    };

    win.onResizing = win.onResize = function () {
        this.layout.resize();
        if (activeTabController) activeTabController.refreshPreview(); // re-fit the active tab's preview box/image
    };

    updateContextBar();
    updateFooterButtons();
}

//----------------------------------------------------------------------
// ENTRY POINT
//----------------------------------------------------------------------

function main() {
    try {
        buildUI();
        setStatus("Ready.", false);

        win.center();
        win.show();
    } catch (error) {
        alert(
            "Sizing and Structure Panel failed to open:\n" + errorMessage(error) +
            (error && error.line ? "\n(line " + error.line + ")" : "")
        );
    }
}

main();
