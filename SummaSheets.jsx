function chooseFolder(promptText) {
    var folder = Folder.selectDialog(promptText);
    if (!folder) {
        alert("לא נבחרה תיקייה עבור: " + promptText);
        throw new Error("לא נבחרה תיקייה");
    }
    return folder;
}

function chooseJSONFile(promptText) {
    var file = File.openDialog(promptText, "*.json");
    if (!file) {
        alert("לא נבחר קובץ JSON!");
        throw new Error("לא נבחר קובץ JSON");
    }
    return file;
}

function cleanString(str) {
    str = str.replace(/[\u200f\u200e]/g, ""); // תווי כיוון
    str = str.replace(/^\s+|\s+$/g, ""); // רווחים מהתחלה וסוף
    str = str.replace(/^["']+|["']+$/g, ""); // גרשיים
    return str;
}

function parseJSON(content) {
    content = content.replace(/\r?\n|\r/g, "");
    var match = content.match(/"files"\s*:\s*\[\s*([^\]]+)\s*\]/);
    if (!match) {
        throw new Error("לא נמצאה רשימת files בקובץ");
    }

    var rawItems = match[1].split(",");
    var items = [];

    for (var i = 0; i < rawItems.length; i++) {
        var raw = rawItems[i];
        var item = cleanString(raw);
        items.push(item);
    }

    return items;
}

function readJSONFile(file) {
    file.open("r");
    var content = file.read();
    file.close();
    return parseJSON(content);
}

// מיון מסגרות: מלמעלה למטה, משמאל לימין
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

// תאריך
var today = new Date();
var day = today.getDate();
var month = today.getMonth() + 1;
if (day < 10) day = "0" + day;
if (month < 10) month = "0" + month;
var dateStr = day + "-" + month;

// שלב 1: בחירת תיקיות
var templateFolder = chooseFolder("בחרי את תיקיית קובץ הטמפלט");
var somaFolder = chooseFolder("בחרי את תיקיית קבצי הסומה (pdf)");
var outputFolder = chooseFolder("בחרי את תיקיית השמירה לקבצים החדשים");
var jsonFile = chooseJSONFile("בחרי את קובץ ה־JSON עם הסדר");

// שלב 2: קריאת JSON
var fileOrder = readJSONFile(jsonFile);
if (!fileOrder || fileOrder.length === 0) {
    alert("קובץ JSON לא מכיל רשימת קבצים תקינה");
    throw new Error("קובץ JSON ריק או לא תקני");
}

// שלב 3: טעינת קבצי סומה לפי שמות
var somaFiles = somaFolder.getFiles(function (f) {
    return f instanceof File && f.name.match(/\.ai$/i);
});

function findSomaFileByName(namePart) {
    namePart = cleanString(namePart);
    for (var i = 0; i < somaFiles.length; i++) {
        var nameNoExt = decodeURI(somaFiles[i].name.replace(/\.ai$/i, ""));
        nameNoExt = cleanString(nameNoExt);
        if (nameNoExt === namePart) {
            return somaFiles[i];
        }
    }
    return null;
}

// פתיחת טמפלט
var templateFile = templateFolder.getFiles("*.idml")[0];
if (!templateFile) {
    alert("לא נמצא קובץ INDD בתיקיית הטמפלט!");
    throw new Error("אין טמפלט");
}

// שלב 4: יצירת הגליונות
var currentIndex = 0;
var batch = 0;
var doc;

while (currentIndex < fileOrder.length) {
    try {
        doc = app.open(templateFile);
    } catch (e) {
        alert("שגיאה בפתיחת הטמפלט: " + e);
        throw e;
    }

    var frames = getSortedFrames(doc);
    var usedCount = Math.min(frames.length, fileOrder.length - currentIndex);
    for (var j = 0; j < usedCount; j++, currentIndex++) {
        var descriptor = fileOrder[currentIndex];
        var fileToPlace = findSomaFileByName(descriptor);
        if (!fileToPlace) {
            alert("לא נמצא קובץ עבור: " + descriptor);
            continue;
        }
        try {
            var frameIndex = frames.length - 1 - j; // מסגרת מהסוף להתחלה

            // app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_ART;//הטמעה לפי גודל הartboard
            // var placed = frames[frameIndex].place(fileToPlace);
            // if (placed) placed[0].itemLink.embed();

            var frame = frames[frameIndex];
            // app.pdfPlacePreferences.createStaticCaptions = true;
            app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_TRIM; //ממקם ביחס לחיתוך הדפסה

            var page = frame.parentPage;
            var placed = frame.place(fileToPlace);

            // var placed = page.place(fileToPlace, [frame.geometricBounds[1], frame.geometricBounds[0]]);
            if (placed && placed.length > 0) {
                var pdf = placed[0];

                // מטמיע את הקובץ
                pdf.itemLink.embed();
            }
        } catch (e) {
            // alert("שגיאה בהנחת קובץ " + fileToPlace.name + ": " + e);
        }
    }

    // עדכון כותרת
    var texts = doc.textFrames;
    var name = "GLIX_סומה " + (batch + 1) + " תאריך " + dateStr;
    for (var i = 0; i < texts.length; i++) {
        var tf = texts[i];
        if (tf.contents.indexOf("גיליון מדבקות") !== -1) {
            tf.contents = name;
            break;
        }
    }

    // שמירה
    var fileName = "GLIX_" + (batch + 1) + " תאריך " + dateStr + ".indd";
    var saveFile = new File(outputFolder + "/" + fileName);
    var pdfFile = new File(outputFolder + "/GLIX_סומה_" + (batch + 1) + " תאריך " + dateStr + ".pdf");

    doc.save(saveFile, true);
    doc.exportFile(ExportFormat.PDF_TYPE, pdfFile, false);
    batch++;
    doc.close(SaveOptions.YES);
}

alert("הסקריפט הסתיים בהצלחה. נוצרו " + batch + " גליונות.");