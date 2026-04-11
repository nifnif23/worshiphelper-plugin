using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace WorshipHelperVSTO
{
    public class Bible
    {
        public string name { get; set; }
        public List<Book> books = new List<Book>();
    }
    public class Book
    {
        public string name { get; set; }
        public List<Chapter> chapters = new List<Chapter>();
    }
    public class Chapter
    {
        public int number { get; set; }
        public List<Verse> verses = new List<Verse>();
    }
    public class Verse
    {
        public int number { get; set; }
        public string text { get; set; }
    }

    public class OpenSongBibleReader
    {
        public static Bible LoadTranslation(String translationName)
        {
            var bible = new OpenSongBibleReader().load($@"{ThisAddIn.appDataPath}\Bibles\{translationName}.xmm");
            bible.name = translationName;
            return bible;
        }

        public Bible load(String fileName)
        {
            var xml = XDocument.Load(fileName);

            var bible = new Bible();

            var bookElements = from item in xml.Descendants("b") select item;
            foreach (XElement bookElement in bookElements)
            {
                var book = new Book();
                book.name = bookElement.Attribute("n").Value;
                bible.books.Add(book);

                var chapterElements = from item in bookElement.Descendants("c") select item;
                foreach (XElement chapterElement in chapterElements)
                {
                    var chapter = new Chapter();
                    chapter.number = Int32.Parse(chapterElement.Attribute("n").Value);
                    book.chapters.Add(chapter);

                    var verseElements = from item in chapterElement.Descendants("v") select item;
                    foreach (XElement verseElement in verseElements)
                    {
                        var verse = new Verse();
                        verse.number = Int32.Parse(verseElement.Attribute("n").Value);
                        verse.text = verseElement.Value;
                        chapter.verses.Add(verse);
                    }
                }
            }

            return bible;
        }
    }
}
