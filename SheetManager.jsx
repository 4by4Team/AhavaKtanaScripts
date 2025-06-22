// פונקציה לפתיחת תיקייה דרך דיאלוג
function chooseFolder(promptText) {
    var folder = Folder.selectDialog(promptText);
    if (!folder) {
        alert("לא נבחרה תיקייה עבור: " + promptText);
        throw new Error("לא נבחרה תיקייה");
    }
    return folder;
}
// קבלת תאריך נוכחי
var today = new Date();
var day = today.getDate();
var month = today.getMonth() + 1; // חודשים מתחילים מ-0

// הוספת אפס מוביל אם צריך
if (day < 10) day = "0" + day;
if (month < 10) month = "0" + month;

// יצירת מחרוזת תאריך
var dateStr = day + "-" + month;

// שלב 1: בחירת התיקיות
var templateFolder = chooseFolder("בחרי את תיקיית קובץ הטמפלט");
var stickersFolder = chooseFolder("בחרי את תיקיית קבצי המדבקות (AI)");
var outputFolder = chooseFolder("בחרי את תיקיית השמירה לקבצים החדשים");

// שלב 2: פתיחת קובץ טמפלט
var templateFile = templateFolder.getFiles("*.idml")[0];
if (!templateFile) {
    alert("לא נמצא קובץ INDD בתיקיית הטמפלט!");
    throw new Error("אין טמפלט");
}

var stickerFiles = stickersFolder.getFiles(function(f) {
    return f instanceof File && f.name.match(/\.ai$/i);
});
stickerFiles.sort(); // לשמור על סדר קבוע

var currentIndex = 0;
var batch = 0;
var doc;

while (currentIndex < stickerFiles.length) {
    // פותח טמפלט חדש עבור עוד 24
    try {
        doc = app.open(templateFile);
    } catch (e) {
        alert("שגיאה בפתיחת הטמפלט: " + e);
        throw e;
    }

    var frames = doc.pageItems; // כל האובייקטים בשכבה "compound path"

    // שלב 6: הוספת קבצי המדבקות לתוך התיבות לפי הסדר
    for (var j = 0; j <= 24 && currentIndex < stickerFiles.length; j++, currentIndex++) {
        try {
            var placed = frames[j].place(stickerFiles[currentIndex]);
            if (!placed) {
                alert("הקובץ לא הונח בהצלחה: " + stickerFiles[currentIndex].name);
            }
        } catch (e) {
            alert("שגיאה בהנחת קובץ " + stickerFiles[currentIndex].name + ": " + e);
        }
    }

    var texts = doc.textFrames; // כל תיבות הטקסט במסמך
    var name = "גליון מדבקות  " + (batch + 1) + "בתאריך   " + dateStr;
    for (var i = 0; i < texts.length; i++) {
        var tf = texts[i];

        // בדיקה אם התיבה מכילה את המילים "גיליון מדבקות"
        if (tf.contents.indexOf("גיליון מדבקות") !== -1) {
            tf.contents = name;
            break;
        }
    }

    // שלב 7: שמירה
    var fileName = "גליון מדבקות_" + (batch + 1) + " בתאריך  " + dateStr + ".indd";
    var saveFile = new File(outputFolder + "/" + fileName);
    var pdfFile = new File(outputFolder + "/גליון מדבקות_" + (batch + 1) + "בתאריך  " + dateStr + ".pdf");
    doc.save(saveFile);
    doc.exportFile(ExportFormat.PDF_TYPE, pdfFile, false);
    batch++;
    doc.close(SaveOptions.YES);

}
alert("הסקריפט הסתיים בהצלחה. נוצרו " + batch + " גליונות.");