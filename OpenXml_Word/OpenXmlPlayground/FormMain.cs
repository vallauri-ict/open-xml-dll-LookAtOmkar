﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPlayground;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using System.Xml;
using System.IO;


namespace OpenXmlPlayground
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
            
        }


        private void btnSimpleWordTest_Click(object sender, EventArgs e)
        {
            try
            {
                string filepath = "test.docx";
                string msg = "Hello World!";
                using (WordprocessingDocument doc = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
                {
                    #region Document
                    // Add a main document part. 
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();

                    // Create the document structure and add some text.
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    #endregion

                    #region Style
                    // Define the styles
                    ClsOpenXmlUtilites.AddStyle(mainPart, false, true, true, "MyHeading1", "Titolone", "Verdana", 30, "0000FF");
                    ClsOpenXmlUtilites.AddStyle(mainPart, true, false, false, "MyTypescript", "Macchina da scrivere", "Courier New", 10, "333333");

                    // Add MyHeading1 styled text
                    Paragraph headingPar = ClsOpenXmlUtilites.CreateParagraphWithStyle("MyHeading1", JustificationValues.Center);
                    ClsOpenXmlUtilites.AddTextToParagraph(headingPar, "Titolo con stile applicato");
                    body.AppendChild(headingPar);

                    // Add simple text
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    // String msg contains the text, "Hello, Word!"
                    run.AppendChild(new Text(msg));

                    // Add MyTypescript styled text

                    Paragraph typescriptParagraph = ClsOpenXmlUtilites.CreateParagraphWithStyle("MyTypescript", JustificationValues.Left);
                    ClsOpenXmlUtilites.AddTextToParagraph(typescriptParagraph, "È universalmente riconosciuto che un lettore che osserva il layout di una pagina viene distratto dal contenuto testuale se questo è leggibile. Lo scopo dell’utilizzo del Lorem Ipsum è che offre una normale distribuzione delle lettere (al contrario di quanto avviene se si utilizzano brevi frasi ripetute, ad esempio “testo qui”), apparendo come un normale blocco di testo leggibile. Molti software di impaginazione e di web design utilizzano Lorem Ipsum come testo modello. Molte versioni del testo sono state prodotte negli anni, a volte casualmente, a volte di proposito (ad esempio inserendo passaggi ironici).");
                    body.AppendChild(typescriptParagraph);

                    // Append a paragraph with styles
                    Paragraph newPar = createParagraphWithStyles();
                    body.AppendChild(newPar);
                    #endregion

                    #region Table
                    // Append a table
                    bool[] bolds = { false, true, false, false };
                    bool[] italics = { false, false, false, false };
                    bool[] underlines = { false, false, false, false };
                    string[] texts1 = { "A", "Nice", "Little", "Table" };
                    JustificationValues[] justifications = { JustificationValues.Left, JustificationValues.Left, JustificationValues.Left, JustificationValues.Center };
                    Table myTable = ClsOpenXmlUtilites.createTable(mainPart, bolds, italics, underlines, texts1, justifications, 2, 2, "CC0000");
                    body.Append(myTable);
                    #endregion

                    #region List and Image
                    // Append bullet list
                    string[] texts2 = { "First element", "Second Element", "Third Element" };
                    ClsOpenXmlUtilites.CreateBulletNumberingPart(mainPart);
                    List<Paragraph> bulletList = new List<Paragraph>();
                    ClsOpenXmlUtilites.CreateBulletOrNumberedList(100, 200, bulletList, texts2.Length, texts2);
                    foreach (Paragraph paragraph in bulletList)
                        body.Append(paragraph);

                    // Append numbered list
                    List<Paragraph> numberedList = new List<Paragraph>();
                    ClsOpenXmlUtilites.CreateBulletOrNumberedList(100, 240, numberedList, texts2.Length, texts2, false);
                    foreach (Paragraph paragraph in numberedList)
                        body.Append(paragraph);

                    // Append image
                    ClsOpenXmlUtilites.InsertPicture(doc, "panorama.jpg");
                    #endregion
                }
                MessageBox.Show("Il documento è pronto!");
                Process.Start(filepath);
            }
            catch (Exception)
            {
                MessageBox.Show("Problemi col documento. Se è aperto da un altro programma, chiudilo e riprova...");
            }
        }

        private Paragraph createParagraphWithStyles()
        {
            Paragraph p = new Paragraph();
            // Set the paragraph properties
            ParagraphProperties pp = new ParagraphProperties(new ParagraphStyleId() { Val = "Titolo1" });
            pp.Justification = new Justification() { Val = JustificationValues.Center };
            // Add paragraph properties to your paragraph
            p.Append(pp);

            // Run 1
            Run r1 = new Run();
            Text t1 = new Text("Pellentesque ") { Space = SpaceProcessingModeValues.Preserve };
            // The Space attribute preserve white space before and after your text
            r1.Append(t1);
            p.Append(r1);

            // Run 2 - Bold
            Run r2 = new Run();
            RunProperties rp2 = new RunProperties();
            rp2.Bold = new Bold();
            // Always add properties first
            r2.Append(rp2);
            Text t2 = new Text("commodo ") { Space = SpaceProcessingModeValues.Preserve };
            r2.Append(t2);
            p.Append(r2);

            // Run 3
            Run r3 = new Run();
            Text t3 = new Text("rhoncus ") { Space = SpaceProcessingModeValues.Preserve };
            r3.Append(t3);
            p.Append(r3);

            // Run 4 – Italic
            Run r4 = new Run();
            RunProperties rp4 = new RunProperties();
            rp4.Italic = new Italic();
            // Always add properties first
            r4.Append(rp4);
            Text t4 = new Text("mauris") { Space = SpaceProcessingModeValues.Preserve };
            r4.Append(t4);
            p.Append(r4);

            // Run 5
            Run r5 = new Run();
            Text t5 = new Text(", sit ") { Space = SpaceProcessingModeValues.Preserve };
            r5.Append(t5);
            p.Append(r5);

            // Run 6 – Italic , bold and underlined
            Run r6 = new Run();
            RunProperties rp6 = new RunProperties();
            rp6.Italic = new Italic();
            rp6.Bold = new Bold();
            rp6.Underline = new Underline() { Val = UnderlineValues.WavyDouble };
            // Always add properties first
            r6.Append(rp6);
            Text t6 = new Text("amet ") { Space = SpaceProcessingModeValues.Preserve };
            r6.Append(t6);
            p.Append(r6);

            // Run 7
            Run r7 = new Run();
            Text t7 = new Text("faucibus arcu ") { Space = SpaceProcessingModeValues.Preserve };
            r7.Append(t7);
            p.Append(r7);

            // Run 8 – Red color
            Run r8 = new Run();
            RunProperties rp8 = new RunProperties();
            rp8.Color = new Color() { Val = "FF0000" };
            // Always add properties first
            r8.Append(rp8);
            Text t8 = new Text("porttitor ") { Space = SpaceProcessingModeValues.Preserve };
            r8.Append(t8);
            p.Append(r8);

            // Run 9
            Run r9 = new Run();
            Text t9 = new Text("pharetra. Maecenas quis erat quis eros iaculis placerat ut at mauris.") { Space = SpaceProcessingModeValues.Preserve };
            r9.Append(t9);
            p.Append(r9);

            // return the new paragraph
            return p;
        }
    }
}
