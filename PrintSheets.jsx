// ניקוי שם קובץ מתווים לא נראים ורווחים
function cleanFileName(name) {
    return name.replace(/\u200E|\u200F/g, "").replace(/^\s+|\s+$/g, "");
}

// שליפת מספר ההזמנה מתוך [#123456]
function extractOrderNumber(name) {
    var clean = cleanFileName(name); // <<< שונה
    var decoded = decodeURIComponent(clean); // <<< שונה
    var match = decoded.match(/\[\s*#?(\d+)\s*\]/); // <<< שונה: ביטוי רגולרי פחות קשיח
    return match ? match[1] : "zzz";
}

// שליפת תגית מדבקה מתוך $...$
function extractStickerLabel(fileName) {
    var clean = cleanFileName(fileName);
    var decoded = decodeURIComponent(clean);
    var match = decoded.match(/\$(.*?)\$/);
    return match ? match[1] : "לא ידוע";
}

// שליפת כמות עותקים מתוך (2)
function extractCopiesCount(fileName) {
    var match = fileName.match(/\((\d+)\)/);
    return match ? parseInt(match[1]) : 1;
}

// מיון קבצים לפי מספר הזמנה, שם בסיס, גרסה
function sortStickerFiles(files) {
    return files.sort(function (a, b) {
        var nameA = cleanFileName(a.name);
        var nameB = cleanFileName(b.name);

        var orderA = extractOrderNumber(nameA);
        var orderB = extractOrderNumber(nameB);
        if (orderA !== orderB) return parseInt(orderA) - parseInt(orderB);

        var baseA = nameA.replace(/_\(\d+\)/, "").toLowerCase();
        var baseB = nameB.replace(/_\(\d+\)/, "").toLowerCase();
        if (baseA !== baseB) return baseA.localeCompare(baseB, "he");

        var verA = (nameA.match(/_\((\d+)\)/) || [0, 0])[1];
        var verB = (nameB.match(/_\((\d+)\)/) || [0, 0])[1];
        return parseInt(verA) - parseInt(verB);
    });
}

// מיון מסגרות (למעלה למטה, שמאל לימין)
function getSortedFrames(doc) {
    var frames = [];
    for (var i = 0; i < doc.rectangles.length; i++) {
        var r = doc.rectangles[i];
        if (r.isValid && !r.locked && r.visible) frames.push(r);
    }

    frames.sort(function (a, b) {
        var ay = a.geometricBounds[0], ax = a.geometricBounds[1];
        var by = b.geometricBounds[0], bx = b.geometricBounds[1];
        return Math.abs(ay - by) > 1 ? ay - by : ax - bx;
    });

    return frames;
}

// המרה לפורמט טקסט JSON פשוט
function simpleStringify(obj) {
    var str = '{\n';
    for (var key in obj) {
        if (obj.hasOwnProperty(key)) {
            var arr = obj[key];
            str += '  "' + key + '": [\n';
            for (var i = 0; i < arr.length; i++) {
                str += '    "' + arr[i] + '"';
                if (i < arr.length - 1) str += ',';
                str += '\n';
            }
            str += '  ]\n';
        }
    }
    str += '}';
    return str;
}

// פונקציית בחירת תיקייה
function chooseFolder(promptText) {
    var folder = Folder.selectDialog(promptText);
    if (!folder) {
        alert("לא נבחרה תיקייה עבור: " + promptText);
        throw new Error("לא נבחרה תיקייה");
    }
    return folder;
}

// תאריך
var today = new Date();
var day = today.getDate(), month = today.getMonth() + 1;
if (day < 10) day = "0" + day;
if (month < 10) month = "0" + month;
var dateStr = day + "-" + month;

// בחירת תיקיות
var templateFolder = chooseFolder("בחרי את תיקיית קובץ הטמפלט");
var stickersFolder = chooseFolder("בחרי את תיקיית קבצי המדבקות (pdf)");
var outputFolder = chooseFolder("בחרי את תיקיית השמירה לקבצים החדשים");

// טמפלט
var templateFiles = templateFolder.getFiles(function (f) {
    return f instanceof File && f.name.match(/\.idml$/i);
});
if (!templateFiles || templateFiles.length === 0) {
    alert("לא נמצא קובץ טמפלט בתיקייה!");
    throw new Error("אין טמפלט");
}
// app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_TRIM;//השמה ביחס לחיתוך הדפסה

// קבצי מדבקה גולמיים
var rawStickerFiles = stickersFolder.getFiles(function (f) {
    return f instanceof File && f.name.match(/\.pdf$/i);
});

// מיון
var stickerFiles = sortStickerFiles(rawStickerFiles);

// בניית רשימה מוכפלת לפי כמות עותקים
var expandedStickerFiles = [];
var stickerLabels = [];
var orderNumbers = [];
for (var i = 0; i < stickerFiles.length; i++) {
    var file = stickerFiles[i];
    var count = extractCopiesCount(file.name);
    var label = extractStickerLabel(file.name);
    var order = extractOrderNumber(file.name);
    for (var c = 0; c < count; c++) {
        expandedStickerFiles.push(file);
        stickerLabels.push(label);
        orderNumbers.push(order);
    }
}

// כתיבת JSON
var jsonContent = simpleStringify({ files: stickerLabels });
var jsonFile = new File(outputFolder + "/GLIX_" + dateStr + ".json");
jsonFile.encoding = "UTF-8";
jsonFile.open("w");
jsonFile.write(jsonContent);
jsonFile.close();

// <<< חדש: שמירת JSON עם הזמנות לפי גיליון
var sheetsOrdersMap = {}; // <<< חדש

// יצירת גליונות
var currentIndex = 0;
var batch = 0;
var doc;

while (currentIndex < expandedStickerFiles.length) {
    try {
        doc = app.open(templateFiles[0]);
    } catch (e) {
        alert("שגיאה בפתיחת טמפלט: " + e);
        throw e;
    }

    var usedCount = Math.min(24, expandedStickerFiles.length - currentIndex);
    var frames = getSortedFrames(doc);
    var currentSheetOrders = []; // <<< חדש

    for (var j = 0; j < usedCount; j++, currentIndex++) {
        try {
            var frameIndex = frames.length - 1 - j; // התחלה מהמסגרת האחרונה אחורה
            var placed = frames[frameIndex].place(expandedStickerFiles[currentIndex]);
            if (placed) placed[0].itemLink.embed();
        } catch (e) {
            $.writeln("שגיאה בהנחת קובץ " + expandedStickerFiles[currentIndex].name + ": " + e);
        }

        var order = extractOrderNumber(expandedStickerFiles[currentIndex].name); // <<< חדש
        currentSheetOrders.push(order); // <<< חדש
    }

    var name = "גליקס " + (batch + 1) + " " + dateStr;
    var texts = doc.textFrames;
    for (var i = 0; i < texts.length; i++) {
        if (texts[i].contents.indexOf("גיליון מדבקות") !== -1) {
            texts[i].contents = name;
            break;
        }
    }

    // שמירת PDF
    var pdfFile = new File(outputFolder + "/" + name + ".pdf");
    doc.exportFile(ExportFormat.PDF_TYPE, pdfFile, false);

    // שמירת JPG
    try {
        app.jpegExportPreferences.exportResolution = 300;
        app.jpegExportPreferences.jpegQuality = JPEGOptionsQuality.MAXIMUM;
        app.jpegExportPreferences.jpegRenderingStyle = JPEGOptionsFormat.BASELINE_ENCODING;

        var jpgFile = new File(outputFolder + "/" + name + ".jpg");
        doc.exportFile(ExportFormat.JPG, jpgFile, false);
    } catch (e) {
        alert("שגיאה בשמירת JPEG: " + e);
    } finally {
        doc.close(SaveOptions.NO);
        sheetsOrdersMap[name] = currentSheetOrders;
        batch++;
    }
}

// <<< חדש: כתיבת JSON של מספרי הזמנות לפי גיליון
var sheetsJsonContent = simpleStringify(sheetsOrdersMap); // <<< חדש
var sheetsJsonFile = new File(outputFolder + "/GLIX_sheet_orders_" + dateStr + ".json");
sheetsJsonFile.encoding = "UTF-8";
sheetsJsonFile.open("w");
sheetsJsonFile.write(sheetsJsonContent);
sheetsJsonFile.close(); // <<< חדש

alert("הסקריפט הסתיים בהצלחה. נוצרו " + batch + " גליונות.");