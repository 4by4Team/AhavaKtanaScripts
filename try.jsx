var adjustedGroups = {};
var adjustedGroupsCount = 0;

function recursiveReplaceWithFonts(item) {
    if (item.typename === "GroupItem") {
        if (!item.name) return;
        if (adjustedGroups[item.name]) {
            adjustedGroupsCount++;
            return;
        }
        for (var i = 0; i < item.pageItems.length; i++) {
            recursiveReplaceWithFonts(item.pageItems[i]);
        }
        adjustedGroups[item.name] = true;
    } else if (item.typename === "TextFrame") {
        // אפשר להוסיף פה פעולה על טקסט אם רוצים
    }
}

if (app.documents.length === 0) {
    alert("No open documents");
} else {
    var doc = app.activeDocument;
    for (var i = 0; i < doc.pageItems.length; i++) {
        recursiveReplaceWithFonts(doc.pageItems[i]);
    }
    alert("Number of skipped groups due to naming: " + adjustedGroupsCount);
}
