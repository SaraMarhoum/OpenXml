using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;

namespace OpenXmlWordTest
{
    public class Program
    {
        static void Main(string[] args)
        {

            //string filepath = @"C:\Users\smarhoum\Documents\Osmose\Word\Prénom Nom.docx";
            string filepath = "TestDoc2.docx";
            //ReadWordDoc(filepath);
            //ReadWordDocLoopParagraph(filepath);
            ReadWordDocLoopParagraphAndSplitToHTML(filepath);
        }

         
        //Lecture du document word docx 
        public static void ReadWordDoc(string filepath)
        {
        try
        {
            // Open a WordprocessingDocument based on a filepath.
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, false))
        {
            // Assign a reference to the existing document body.
           Body body = wordDocument.MainDocumentPart.Document.Body;
           Console.Write(body.InnerText);
        }
        }
        catch (Exception ex)
        {
            throw ex;
        }
        }


        //Lecture du document word docx par paragraphe
        public static void ReadWordDocLoopParagraph(string filepath)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, false))
            {

                var paragraphs = wordDocument.MainDocumentPart.RootElement.Descendants<Paragraph>();

                foreach (var paragraph in paragraphs)
                {
                    Console.WriteLine(paragraph.InnerText);
                }
                Console.ReadKey();
            }

        }

        //Lecture du document word docx par paragraphe et split to html docs
        public static void ReadWordDocLoopParagraphAndSplitToHTML(string filepath)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, true))
            {
                //var writer=File.CreateText("Paragraph4.txt");
                var Lines = wordDocument.MainDocumentPart.RootElement.Descendants<ParagraphStyleId>();

                var paragraphTextes = wordDocument.MainDocumentPart.RootElement.Descendants<Paragraph>();

                    foreach (var line in Lines) 
                    {
                       
                       if(line.Val != "Titre1")
                       {

                            //Delete the line from the docx
                            
                       }
                         
                    }
            }

        }
    }
}


