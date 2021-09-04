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

            string filepath = "TestDoc2.docx";
            //ReadWordDoc(filepath);
            //ReadWordDocLoopParagraph(filepath);
            //ReadWordDocLoopParagraphAndSplitToHTML(filepath);
            FindstyleParagraphs(filepath);
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

                // string [] parArray;

                foreach (var paragraph in paragraphs)
                {
                    string par = paragraph.InnerText;

                    //parArray[] = par;

                    //Console.WriteLine(paragraph.LocalName());
                }

                // foreach( int i in parArray)
                // {
                //     Console.WriteLine(parArray[i]);

                // }

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

                var paragraphs= wordDocument.MainDocumentPart.RootElement.Descendants<Paragraph>();

                    foreach (var paragraph in paragraphs) 
                    {

                        Console.WriteLine(paragraph.InnerText);
                       
                    //    if(line.Val == "Titre1" || line.Val == "titre2" || line.Val == "titre3" )
                    //    {

                    //         //Delete the line from the docx
                            
                    //    }
                         
                    }
            }

        }

        public static void FindstyleParagraphs(string filepath)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, true))
            {
                var paragraphs = new List<Paragraph>();
                paragraphs = wordDocument.MainDocumentPart.Document.Body
                .OfType<Paragraph>()
                .Where(p => p.ParagraphProperties != null && 
                        p.ParagraphProperties.ParagraphStyleId != null && 
                        p.ParagraphProperties.ParagraphStyleId.Val.Value.Contains("Titre1")).ToList();

                        foreach (var paragraph in paragraphs)
                        {
                            Console.WriteLine(paragraph.InnerText);
                            File.CreateText(paragraph.InnerText + ".txt");
                        }
                            
            }
        }
    }
}


