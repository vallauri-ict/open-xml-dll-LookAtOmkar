#                                              OPEN-XML-DLL
*Quì troviamo un progetto realizzato in C# "OpenXml", in cui attraverso i codici andiamo a creare un file di Word scritto,come facciamo noi normalmente sull'app di microsoft office.*
*Abbiamo usato DLL(Dynamic Link Library) per:*
*  semplificare la lettura dei codici presenti nel progetto
*  la programmazione senza complicità
*  la sua realizzazione con comodità.
**OPEN-XML-WORD**
*nella DLL del progetto OpenXml di Word contengono:*
* public static void InsertPicture(WordprocessingDocument wordprocessingDocument, string fileName){...}
* private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId){...}
* public static void AddStyle(MainDocumentPart mainPart, eventuali parametri){...}
* public static Paragraph CreateParagraphWithStyle(string styleId, JustificationValues justification      JustificationValues.Left){...}
* public static void AddTextToParagraph(Paragraph paragraph, string content){...}
* public static void InsertParagraphInList(List<Paragraph> paragraphs, ParagraphProperties ppUnordered, string text){...}
*Poi ci sono altri, ma diciamo che questi sono abbastanza importanti per quanto riguarda Word*

**OPEN-XML-EXCEL**
*Nella DLL del progetto OpenXml di Excel contengono:*
* public static void CreateExcelFile(TestModelList data, string OutPutFileDirectory){...}
* private static void CreatePartsForExcel(SpreadsheetDocument document, TestModelList data){...}
* private static void GenerateWorkbookPartContent(WorkbookPart workbookPart1){...}
* private static void GenerateWorksheetPartContent(WorksheetPart worksheetPart1, SheetData sheetData1){...}
* private static void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart1){...}
* private static Row CreateHeaderRowForExcel(){...}
* private static Cell CreateCell(string text){...}
*Poi ci sono altri *
*Con questi DLL, riusciamo a gestire al meglio la programmazione, dato che sono abbastanza complicati a gestirli  in generale specialmente quando si tratta di Microsoft Office*  

**OMKAR SINGH RATHORE 4B INF**


***open-xml-dll-LookAtOmkar created by GitHub Classroom***
