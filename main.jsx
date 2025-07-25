#target illustrator
#include "json2.js"

function safeTrim(str) {
    return (str && typeof str === "string") ? str.replace(/^\s+|\s+$/g, "") : "";
}
function loadFontMapFromSameFolder(filename) {
    var scriptFolder = File($.fileName).parent;
    var fontFile = new File(scriptFolder + "/" + filename);

    if (!fontFile.exists) {
        alert("⚠️ קובץ פונטים לא נמצא: " + fontFile.fsName);
        return;
    }

    try {
        fontFile.open("r");
        var content = fontFile.read();
        fontFile.close();
        var fontList = JSON.parse(content);

        var result = {};
        for (var i = 0; i < fontList.length; i++) {
            var item = fontList[i];
            var modelName = safeTrim(item["דגם"]);
            result[modelName] = {
                hebrew: safeTrim(item["פונט עברית"]),
                english: safeTrim(item["פונט אנגלית"])
            };
        }
        return result;
    } catch (e) {
        alert("⚠️ שגיאה בקריאת קובץ פונטים: " + e.message);
    }
    return {};
}
function fitPointTextToFrame(textFrame, minSize, maxSize, maxWidth, maxHeight) {
    if (!textFrame || !textFrame.textRange || textFrame.kind !== TextType.POINTTEXT) {
        alert("❌ לא מדובר ב-PointText תקין");
        return;
    }

    var textRange = textFrame.textRange;
    var currentSize = textRange.characterAttributes.size;
    var bestSize = currentSize;
    var low = minSize;
    var high = maxSize;
    var tolerance = 0.8;
    var steps = 0;
    var maxSteps = 20;

    var bounds = textFrame.visibleBounds;
    var frameHeight = bounds[0] - bounds[2];
    var frameWidth = bounds[3] - bounds[1];
    if (frameWidth > maxWidth + tolerance || frameHeight > maxHeight + tolerance) {
        while (high - low > 0.1 && steps < maxSteps) {
            var mid = (low + high) / 2;
            textRange.characterAttributes.size = mid;
            app.redraw();

            bounds = textFrame.visibleBounds;
            frameHeight = bounds[0] - bounds[2];
            frameWidth = bounds[3] - bounds[1];

            if (frameWidth <= maxWidth + tolerance && frameHeight <= maxHeight + tolerance) {
                bestSize = mid;
                low = mid;
            } else {
                high = mid;
            }

            steps++;
        }
    }

    textRange.characterAttributes.size = bestSize;
    return bestSize;
}

function splitTextSmart(text) {
    if (typeof text !== "string") {
        throw new Error("splitTextSmart קיבלה ערך שאינו מחרוזת: " + text);
    }
    var words = text.split(" ");
    if (words.length <= 1) return "\u200C\r" + words[0];
    if (words.length === 2) return words[0] + "\r" + words[1];

    var bestSplit = "";
    var minDiff = Infinity;
    for (var i = 1; i < words.length; i++) {
        var line1 = words.slice(0, i).join(" ");
        var line2 = words.slice(i).join(" ");
        var diff = Math.abs(line1.length - line2.length);
        if (diff < minDiff) {
            minDiff = diff;
            bestSplit = line1 + "\r" + line2;
        }
    }
    return bestSplit;
}

function centerTextVertically(textFrame, targetHeight) {
    var delta = (targetHeight - textFrame.height) / 2;
    textFrame.top -= delta;
}

function countWords(str) {
    return str.replace(/^\s+|\s+$/g, "").split(/\s+/).length;
}

function replaceAllTextinDoc(doc, ordernumber, name, minSize, maxSize, fonts) {
    if (!doc) return;
    var smartSplitName = splitTextSmart(name);
    var tf = null;
    for (var i = 0; i < doc.textFrames.length; i++) {
        if (doc.textFrames[i].contents.indexOf("0000000") !== -1) {
            tf = doc.textFrames[i];
            break;
        }
    }
    if (tf) {
        tf.contents = tf.contents.replace(/0000000/g, ordernumber);
    }
    for (var i = 0; i < doc.pageItems.length; i++) {
        recursiveReplaceWithFonts(doc.pageItems[i], "שמות הילד", name, minSize, maxSize, fonts);
        recursiveReplaceWithFonts(doc.pageItems[i], "שמות\rהילד", smartSplitName, minSize, maxSize, fonts);
    }
}

function recursiveReplaceWithFonts(item, searchText, replaceText, minSize, maxSize, fonts) {
    if (item.typename === "GroupItem") {
        if (!item.name) return;
        for (var i = 0; i < item.pageItems.length; i++) {
            recursiveReplaceWithFonts(item.pageItems[i], searchText, replaceText, minSize, maxSize, fonts);
        }
    } else if (item.typename === "TextFrame") {
        var tfText = item.contents;
        var isMultiline = tfText.indexOf("\r") >= 0;
        var pattern = searchText.replace(/\s+/g, isMultiline ? "\\s*\\r?\\s*" : "\\s+");
        var regex = new RegExp(pattern, "g");
        if (regex.test(tfText)) {
            cleanAndReplaceWithFonts(item, searchText, replaceText, minSize, maxSize, fonts);
        }
    }
}

function cleanAndReplaceWithFonts(tf, original, replacement, minSize, maxSize, fonts) {
    if (!tf || !tf.textRange) {
        return;
    }

    // var isMultiline = tf.contents.indexOf("\r") >= 0;
    // var newText = isMultiline ? splitTextSmart(replacement) : replacement;
    var isMultiline = original.indexOf("\r") >= 0;
    var newText = isMultiline ? splitTextSmart(replacement) : replacement;
    var pattern = original.replace(/\s+/g, isMultiline ? "\\s*\\r?\\s*" : "\\s+");
    var regex = new RegExp(pattern, "g");
    var fullRange = tf.textRange;
    var originalSize = fullRange.characterAttributes.size;
    var originalColor = fullRange.characterAttributes.fillColor;
    var replaced = tf.contents.replace(regex, newText);
    if (replaced === tf.contents) {
        return;
    }

    tf.contents = replaced;
    var originalWords = countWords(original);
    var replacementWords = countWords(replacement);
    var shrinkFactor = 1;

    if (replacementWords >= 4) {
        shrinkFactor = minSize / originalSize;
    } else if (replacementWords > originalWords) {
        shrinkFactor = 0.9;
    } else if (replacementWords < originalWords) {
        shrinkFactor = 1.1;
    }

    var range = tf.textRange;

    // קביעת הפונט לפי שפת הטקסט (עברית / אנגלית)
    var isHebrew = /[\u0590-\u05FF]/.test(newText);
    var fontName = isHebrew ? fonts.hebrew : fonts.english;
    var textFont;
    try {
        textFont = app.textFonts.getByName(fontName);
    } catch (e) {
        textFont = range.characterAttributes.textFont; // משאיר כמו שהיה
    }

    range.characterAttributes.textFont = textFont;
    range.characterAttributes.size = originalSize * shrinkFactor;
    range.characterAttributes.fillColor = originalColor;

    var maxWidth = tf.width;
    var maxHeight = tf.height;
    app.redraw();

    fitPointTextToFrame(tf, minSize, maxSize, maxWidth, maxHeight);
    centerTextVertically(tf, maxHeight); {}
}

function generateStickersPages(data, templateFolder, outputFolder) {
    for (var i = 0; i < data.length; i++) {
        var item = data[i];

        // בדיקת שדות חובה
        if (!item.name || !item.itemName || !item.referenceNumber || !item.quantity) {
            alert("❌ נתונים חסרים בשורה " + (i + 1));
            continue;
        }

        // יצירת נתיב לקובץ תבנית
        var templatePath = templateFolder.fsName + "\\" + item.itemName + ".ai";
        var templateFile = new File(templatePath);

        if (!templateFile.exists) {
            alert("❌ לא נמצאה תבנית עבור '" + item.itemName + "' בשורה " + (i + 1));
            continue;
        }

        // פתיחת התבנית
        var doc;
        try {
            doc = app.open(templateFile);
        } catch (e) {
            alert("❌ שגיאה בפתיחת הקובץ '" + templateFile.name + "' בשורה " + (i + 1));
            continue;
        }

        // קביעת פונטים לפי המפה
        var cleanName = item.itemName.replace(/^\s+|\s+$/g, '');

        var fonts = null;

        for (var key in fontMap) {
            if (cleanName.indexOf(key) === 0) {
                fonts = fontMap[key];
                break;
            }
        }

        if (!fonts) {
            fonts = {
                hebrew: "ArialHebrew",
                english: "Arial"
            };
        }


        // החלפת טקסט
        try {
            replaceAllTextinDoc(doc, item.referenceNumber, item.name, 6, 15, fonts);
        } catch (e) {
            alert("❌ שגיאה בהחלפת טקסט עבור '" + item.name + "' בשורה " + (i + 1));
            doc.close(SaveOptions.DONOTSAVECHANGES);
            continue;
        }

        // שם הקובץ כולל הכמות
        var baseName = item.referenceNumber + " (" + item.quantity + ").pdf";
        var baseFullPath = outputFolder.fsName + "\\" + baseName;

        // הגדרות PDF
        var pdfOptions = new PDFSaveOptions();
        pdfOptions.pDFPreset = "[Illustrator Default]";

        // שמירה
        try {
            doc.saveAs(new File(baseFullPath), pdfOptions);
        } catch (e) {
            alert("❌ שגיאה בשמירת PDF עבור '" + item.referenceNumber + "': " + e.message);
            doc.close(SaveOptions.DONOTSAVECHANGES);
            continue;
        }

        // סגירה
        try {
            doc.close(SaveOptions.DONOTSAVECHANGES);
        } catch (e) {
            alert("⚠️ שגיאה בסגירת הקובץ עבור '" + item.referenceNumber + "': " + e.message);
        }
    }

    alert("✅ יצירת המדבקות הסתיימה בהצלחה.");
}


var win = new Window("dialog", "יצירת דפי מדבקות");
win.orientation = "column";
win.alignChildren = ["fill", "top"];

function addFolderRow(labelText) {
    var group = win.add("group");
    group.add("statictext", undefined, labelText);
    var pathField = group.add("edittext", undefined, "", {
        readonly: true
    });
    pathField.characters = 40;
    var browseBtn = group.add("button", undefined, "בחר");
    return {
        pathField: pathField,
        browseBtn: browseBtn
    };
}

var jsonRow = addFolderRow("תיקיית קבצי JSON:");
var templatesRow = addFolderRow("תיקיית תבניות:");
var outputRow = addFolderRow("תיקיית יעד:");
var fontMap = loadFontMapFromSameFolder("fonts.json");
var selectedJSONFolder = null;
var selectedTemplateFolder = null;
var selectedOutputFolder = null;

jsonRow.browseBtn.onClick = function() {
    selectedJSONFolder = Folder.selectDialog("בחר תיקיית קבצי JSON");
    if (selectedJSONFolder) jsonRow.pathField.text = selectedJSONFolder.fsName;
};
templatesRow.browseBtn.onClick = function() {
    selectedTemplateFolder = Folder.selectDialog("בחר תיקיית תבניות");
    if (selectedTemplateFolder) templatesRow.pathField.text = selectedTemplateFolder.fsName;
};
outputRow.browseBtn.onClick = function() {
    selectedOutputFolder = Folder.selectDialog("בחר תיקיית יעד לשמירה");
    if (selectedOutputFolder) outputRow.pathField.text = selectedOutputFolder.fsName;
};

var runBtn = win.add("button", undefined, "הפעל סקריפט");
runBtn.onClick = function() {
    if (!selectedJSONFolder || !selectedTemplateFolder || !selectedOutputFolder) {
        alert("יש לבחור את כל שלוש התיקיות לפני הרצה.");
        return;
    }

    var jsonFiles = selectedJSONFolder.getFiles("*.json");
    if (jsonFiles.length === 0) {
        alert("לא נמצאו קבצי JSON בתיקייה.");
        return;
    }

    win.close();

    for (var i = 0; i < jsonFiles.length; i++) {
        var file = jsonFiles[i];

        if (i > 0) {
            var confirmWin = new Window("dialog", "אישור מעבר לקובץ הבא");
            confirmWin.add("statictext", undefined, "לעבד את הקובץ הבא?\n" + file.name);
            var btnGroup = confirmWin.add("group");
            var yesBtn = btnGroup.add("button", undefined, "כן");
            var noBtn = btnGroup.add("button", undefined, "דלג");
            var shouldContinue = false;

            yesBtn.onClick = function() {
                shouldContinue = true;
                confirmWin.close();
            };
            noBtn.onClick = function() {
                shouldContinue = false;
                confirmWin.close();
            };

            confirmWin.center();
            confirmWin.show();

            if (!shouldContinue) continue;
        }

        var jsonContent, data;
        try {
            file.open("r");
            jsonContent = file.read();
            file.close();
            data = JSON.parse(jsonContent);
        } catch (e) {
            alert("⚠️ שגיאה בקריאת קובץ JSON: " + file.name + "\n" + e.message);
            continue;
        }

        try {
            generateStickersPages(data, selectedTemplateFolder, selectedOutputFolder);
        } catch (e) {
            alert("⚠️ שגיאה בהרצת הקובץ '" + file.name + "': " + e.message);
            continue;
        }
    }

    alert("✅ כל קבצי ה־JSON עובדו בהצלחה.");
};

win.center();
win.show();