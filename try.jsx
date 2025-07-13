var doc = app.activeDocument;

var pdfFile = new File("~/Desktop/ExportedFile.pdf");

var pdfSaveOpts = new PDFSaveOptions();
pdfSaveOpts.preserveEditability = true; // שומר שכבות ועריכה  
// pdfSaveOpts.compatibility = PDFCompatibility.ACROBAT_6; // אפשר לשנות לפי צורך  
pdfSaveOpts.generateThumbnails = true;  
pdfSaveOpts.optimization = true;

doc.saveAs(pdfFile, pdfSaveOpts);