using System.Collections.Generic;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Net;
using System;
using Microsoft.VisualBasic.FileIO;

namespace csv2pdf
{
    /// <summary>
    /// This program is used as a hack to read a CSV of data then using a PDF as a template, create a new PDF with the overlayed data.
    /// </summary>
    class Program
    {
        readonly static string sourcePDF = "../../files/MakerFaire_ExhibitSheet.pdf";
        static void Main(string[] args)
        {
            // Parse the csv to get the required data
            List<AppData> data = parseCsv(args[0]);
            //For each element, create a new PDF, add the text, then save a new copy
            data.ForEach(d => addTextToPdf(d));
        }

        /// <summary>
        /// AppData contains the information needed to create a new PDF with overlayed information including name, description, website, and image.
        /// </summary>
        /// <param name="data">The data to use</param>
        static void addTextToPdf(AppData data)
        {
            using (PdfStamper stamper = new PdfStamper(new PdfReader(sourcePDF), File.Create(data.filename)))
            {
                // Create the header textbox
                TextField name = new TextField(stamper.Writer, new Rectangle(80, 932, 719, 1022), "Name");
                name.Text = data.name;
                name.Alignment = Element.ALIGN_CENTER;

                // Get and set the image
                Image image;
                try
                {
                    image = Image.GetInstance(data.photo);
                } catch
                {
                    image = Image.GetInstance("../../files/default.png");
                }
                
                image.ScaleAbsolute(new Rectangle(80, 200, 721, 593));
                image.SetAbsolutePosition(79, 488);

                // Create the description textbox
                TextField description = new TextField(stamper.Writer, new Rectangle(80, 74, 719, 438), "Description");
                description.Options = TextField.MULTILINE;
                description.Text = data.description;
                description.Alignment = Element.ALIGN_CENTER;

                //calculate the font size
                float size = ColumnText.FitText(new Font(description.Font), data.description, new Rectangle(0, 0, 560, 365), 80, 0);
                description.FontSize = size;

                //Add the website URL if it exists
                TextField website = new TextField(stamper.Writer, new Rectangle(80, 30, 719, 60), "Website");
                website.Text = data.website;
                website.Alignment = Element.ALIGN_CENTER;

                // Write and close
                stamper.AddAnnotation(name.GetTextField(), 1);
                stamper.GetOverContent(1).AddImage(image);
                stamper.AddAnnotation(description.GetTextField(), 1);
                stamper.AddAnnotation(website.GetTextField(), 1);
                stamper.Close();
            }

            
        }

        /// <summary>
        /// Parse the data crv that contains 4 columns: name, description, photo url, and website url
        /// </summary>
        /// <param name="filename">The filename of the CSV to parse</param>
        /// <returns>A list of AppData elements</returns>
        static List<AppData> parseCsv(string filename)
        {
            //Create a new list that will contain the data we parsed
            var data = new List<AppData>();
            //Open the parser
            using (TextFieldParser parser = new TextFieldParser(filename))
            {
                //Setup the parser
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                
                //Loop through each row
                while (!parser.EndOfData)
                {
                    //Processing row
                    string[] fields = parser.ReadFields();
                    data.Add(new AppData(fields[0], fields[1], fields[2], fields[3]));
                }
            }
            return data;
        }
    }

    public class AppData
    {
        public string name { get;}
        public string description { get; }
        public string photo { get; }
        public string filename { get; }
        public string website { get; }

        public AppData(string name, string description, string website, string photo)
        {
            this.name = name.Trim();
            string tmpName = this.name.Replace(" ", "_");
            tmpName = tmpName.Replace(":", "");
            tmpName = tmpName.Replace(".", "");
            tmpName = tmpName.Replace("-", "");
            tmpName = tmpName.Replace("\"", "");
            tmpName = tmpName.Replace("/", "");
            this.filename = "../.../output/" + tmpName + "_ExhibitSheet.pdf";
            this.description = description.Trim();
            //If there is no photo, use the default one
            this.photo = string.IsNullOrEmpty(photo) ? this.photo = "../../files/default.png" : photo;
            this.website = website;
        }
    }
}
