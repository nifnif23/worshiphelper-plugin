using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace WorshipHelperVSTO
{
    class ScriptureReference
    {
        String bookName;
        int chapterNum;
        int verseNumStart;
        int verseNumEnd;

        public static ScriptureReference parse(Bible bible, String bookName, String reference)
        {
            var scriptureReference = new ScriptureReference();

            var book = bible.books.Find(bookItem => bookItem.name.ToLower() == bookName.ToLower());
            if (book == null) throw new Exception("Book does not exist");

            var referenceParts = reference.Split(new char[] { ':', '-' });

            scriptureReference.bookName = book.name;
            scriptureReference.chapterNum = Int32.Parse(referenceParts[0]);
            var chapter = book.chapters.Find(chapterItem => chapterItem.number == scriptureReference.chapterNum);
            if (chapter == null) throw new Exception("Chapter does not exist");

            if (referenceParts.Length > 2)
            {
                scriptureReference.verseNumStart = Int32.Parse(referenceParts[1]);
                scriptureReference.verseNumEnd = Int32.Parse(referenceParts[2]);
            }
            else if (referenceParts.Length > 1)
            {
                scriptureReference.verseNumStart = Int32.Parse(referenceParts[1]);
                scriptureReference.verseNumEnd = scriptureReference.verseNumStart;
            }
            else
            {
                // No verses were specified, so use the whole range
                scriptureReference.verseNumStart = 1;
                scriptureReference.verseNumEnd = chapter.verses.Last().number;
            }

            if (scriptureReference.verseNumEnd < scriptureReference.verseNumStart) throw new Exception("Verse range end is before start");
            if (scriptureReference.verseNumStart < chapter.verses.First().number)  throw new Exception("Verse range is before beginning of chapter");
            if (scriptureReference.verseNumEnd > chapter.verses.Last().number) throw new Exception("Verse range goes past end of chapter");

            return scriptureReference;
        }
    }

    // -----------------------------------------------------------------------
    // Full-reference parser: accepts "Book Chapter:Verse" in one string.
    // Useful for bulk-paste mode where the user types everything in one go.
    // -----------------------------------------------------------------------
    public static class FullReferenceParser
    {
        // Common abbreviation → canonical name map.
        // The keys are stored in lower-case for case-insensitive matching.
        private static readonly Dictionary<string, string> Abbreviations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // Old Testament
            {"gen","Genesis"},{"ge","Genesis"},{"gn","Genesis"},
            {"exo","Exodus"},{"ex","Exodus"},{"exod","Exodus"},
            {"lev","Leviticus"},{"le","Leviticus"},{"lv","Leviticus"},
            {"num","Numbers"},{"nu","Numbers"},{"nm","Numbers"},{"nb","Numbers"},
            {"deut","Deuteronomy"},{"de","Deuteronomy"},{"dt","Deuteronomy"},
            {"josh","Joshua"},{"jos","Joshua"},{"jsh","Joshua"},
            {"judg","Judges"},{"jdg","Judges"},{"jg","Judges"},{"jdgs","Judges"},
            {"ruth","Ruth"},{"rth","Ruth"},{"ru","Ruth"},
            {"1sam","1 Samuel"},{"1sa","1 Samuel"},{"1 sam","1 Samuel"},{"1 sa","1 Samuel"},{"i sam","1 Samuel"},{"i sa","1 Samuel"},
            {"2sam","2 Samuel"},{"2sa","2 Samuel"},{"2 sam","2 Samuel"},{"2 sa","2 Samuel"},{"ii sam","2 Samuel"},{"ii sa","2 Samuel"},
            {"1kgs","1 Kings"},{"1ki","1 Kings"},{"1 kgs","1 Kings"},{"1 ki","1 Kings"},{"i kgs","1 Kings"},{"i ki","1 Kings"},{"1 kings","1 Kings"},
            {"2kgs","2 Kings"},{"2ki","2 Kings"},{"2 kgs","2 Kings"},{"2 ki","2 Kings"},{"ii kgs","2 Kings"},{"ii ki","2 Kings"},{"2 kings","2 Kings"},
            {"1chr","1 Chronicles"},{"1ch","1 Chronicles"},{"1 chr","1 Chronicles"},{"1 ch","1 Chronicles"},{"i chr","1 Chronicles"},{"i ch","1 Chronicles"},{"1 chronicles","1 Chronicles"},
            {"2chr","2 Chronicles"},{"2ch","2 Chronicles"},{"2 chr","2 Chronicles"},{"2 ch","2 Chronicles"},{"ii chr","2 Chronicles"},{"ii ch","2 Chronicles"},{"2 chronicles","2 Chronicles"},
            {"ezra","Ezra"},{"ezr","Ezra"},
            {"neh","Nehemiah"},{"ne","Nehemiah"},
            {"esth","Esther"},{"est","Esther"},{"es","Esther"},
            {"job","Job"},{"jb","Job"},
            {"psa","Psalms"},{"ps","Psalms"},{"psalm","Psalms"},{"pss","Psalms"},{"psm","Psalms"},{"pslm","Psalms"},
            {"prov","Proverbs"},{"pr","Proverbs"},{"prv","Proverbs"},{"pro","Proverbs"},
            {"eccl","Ecclesiastes"},{"ecc","Ecclesiastes"},{"ec","Ecclesiastes"},{"qoh","Ecclesiastes"},
            {"song","Song of Solomon"},{"sos","Song of Solomon"},{"so","Song of Solomon"},{"canticles","Song of Solomon"},{"canticle","Song of Solomon"},{"song of songs","Song of Solomon"},
            {"isa","Isaiah"},{"is","Isaiah"},
            {"jer","Jeremiah"},{"je","Jeremiah"},{"jr","Jeremiah"},
            {"lam","Lamentations"},{"la","Lamentations"},
            {"ezek","Ezekiel"},{"eze","Ezekiel"},{"ezk","Ezekiel"},
            {"dan","Daniel"},{"da","Daniel"},{"dn","Daniel"},
            {"hos","Hosea"},{"ho","Hosea"},
            {"joel","Joel"},{"joe","Joel"},{"jl","Joel"},
            {"amos","Amos"},{"am","Amos"},
            {"obad","Obadiah"},{"ob","Obadiah"},{"oba","Obadiah"},
            {"jonah","Jonah"},{"jon","Jonah"},{"jnh","Jonah"},
            {"mic","Micah"},{"mi","Micah"},
            {"nah","Nahum"},{"na","Nahum"},
            {"hab","Habakkuk"},{"hb","Habakkuk"},
            {"zeph","Zephaniah"},{"zep","Zephaniah"},{"zp","Zephaniah"},
            {"hag","Haggai"},{"hg","Haggai"},
            {"zech","Zechariah"},{"zec","Zechariah"},{"zc","Zechariah"},
            {"mal","Malachi"},{"ml","Malachi"},

            // New Testament
            {"matt","Matthew"},{"mt","Matthew"},{"mat","Matthew"},
            {"mark","Mark"},{"mrk","Mark"},{"mk","Mark"},{"mr","Mark"},
            {"luke","Luke"},{"lk","Luke"},{"lu","Luke"},
            {"john","John"},{"jn","John"},{"jhn","John"},
            {"acts","Acts"},{"act","Acts"},{"ac","Acts"},
            {"rom","Romans"},{"ro","Romans"},{"rm","Romans"},
            {"1cor","1 Corinthians"},{"1co","1 Corinthians"},{"1 cor","1 Corinthians"},{"1 co","1 Corinthians"},{"i cor","1 Corinthians"},{"i co","1 Corinthians"},{"1 corinthians","1 Corinthians"},
            {"2cor","2 Corinthians"},{"2co","2 Corinthians"},{"2 cor","2 Corinthians"},{"2 co","2 Corinthians"},{"ii cor","2 Corinthians"},{"ii co","2 Corinthians"},{"2 corinthians","2 Corinthians"},
            {"gal","Galatians"},{"ga","Galatians"},
            {"eph","Ephesians"},{"ep","Ephesians"},
            {"phil","Philippians"},{"php","Philippians"},{"pp","Philippians"},
            {"col","Colossians"},{"co","Colossians"},
            {"1thess","1 Thessalonians"},{"1th","1 Thessalonians"},{"1 thess","1 Thessalonians"},{"1 th","1 Thessalonians"},{"i thess","1 Thessalonians"},{"i th","1 Thessalonians"},{"1 thessalonians","1 Thessalonians"},
            {"2thess","2 Thessalonians"},{"2th","2 Thessalonians"},{"2 thess","2 Thessalonians"},{"2 th","2 Thessalonians"},{"ii thess","2 Thessalonians"},{"ii th","2 Thessalonians"},{"2 thessalonians","2 Thessalonians"},
            {"1tim","1 Timothy"},{"1ti","1 Timothy"},{"1 tim","1 Timothy"},{"1 ti","1 Timothy"},{"i tim","1 Timothy"},{"i ti","1 Timothy"},{"1 timothy","1 Timothy"},
            {"2tim","2 Timothy"},{"2ti","2 Timothy"},{"2 tim","2 Timothy"},{"2 ti","2 Timothy"},{"ii tim","2 Timothy"},{"ii ti","2 Timothy"},{"2 timothy","2 Timothy"},
            {"titus","Titus"},{"tit","Titus"},{"ti","Titus"},
            {"phlm","Philemon"},{"phm","Philemon"},{"philem","Philemon"},
            {"heb","Hebrews"},{"he","Hebrews"},
            {"james","James"},{"jas","James"},{"jm","James"},
            {"1pet","1 Peter"},{"1pe","1 Peter"},{"1pt","1 Peter"},{"1 pet","1 Peter"},{"1 pe","1 Peter"},{"1 pt","1 Peter"},{"i pet","1 Peter"},{"i pe","1 Peter"},{"i pt","1 Peter"},{"1 peter","1 Peter"},
            {"2pet","2 Peter"},{"2pe","2 Peter"},{"2pt","2 Peter"},{"2 pet","2 Peter"},{"2 pe","2 Peter"},{"2 pt","2 Peter"},{"ii pet","2 Peter"},{"ii pe","2 Peter"},{"ii pt","2 Peter"},{"2 peter","2 Peter"},
            {"1john","1 John"},{"1jn","1 John"},{"1jhn","1 John"},{"1 john","1 John"},{"1 jn","1 John"},{"1 jhn","1 John"},{"i john","1 John"},{"i jn","1 John"},{"i jhn","1 John"},
            {"2john","2 John"},{"2jn","2 John"},{"2jhn","2 John"},{"2 john","2 John"},{"2 jn","2 John"},{"2 jhn","2 John"},{"ii john","2 John"},{"ii jn","2 John"},{"ii jhn","2 John"},
            {"3john","3 John"},{"3jn","3 John"},{"3jhn","3 John"},{"3 john","3 John"},{"3 jn","3 John"},{"3 jhn","3 John"},{"iii john","3 John"},{"iii jn","3 John"},{"iii jhn","3 John"},
            {"jude","Jude"},{"jd","Jude"},{"jud","Jude"},
            {"rev","Revelation"},{"re","Revelation"},{"rv","Revelation"},{"apocalypse","Revelation"},
        };

        /// <summary>
        /// Resolves a user-typed book name (which may be an abbreviation,
        /// partial match, or exact canonical name) to the canonical book
        /// name used in the loaded Bible data.
        /// </summary>
        public static string ResolveBookName(Bible bible, string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;

            string trimmed = input.Trim();

            // 1) Exact match against Bible data (case-insensitive)
            var exact = bible.books.FirstOrDefault(b => b.name.Equals(trimmed, StringComparison.OrdinalIgnoreCase));
            if (exact != null) return exact.name;

            // 2) Abbreviation table
            if (Abbreviations.TryGetValue(trimmed, out string canonical))
            {
                var found = bible.books.FirstOrDefault(b => b.name.Equals(canonical, StringComparison.OrdinalIgnoreCase));
                if (found != null) return found.name;
            }

            // 3) StartsWith match (e.g. "Gene" → "Genesis")
            var startsWith = bible.books.FirstOrDefault(b => b.name.StartsWith(trimmed, StringComparison.OrdinalIgnoreCase));
            if (startsWith != null) return startsWith.name;

            // 4) Contains match (e.g. "Solomon" → "Song of Solomon")
            var contains = bible.books.FirstOrDefault(b => b.name.IndexOf(trimmed, StringComparison.OrdinalIgnoreCase) >= 0);
            if (contains != null) return contains.name;

            return null;
        }

        /// <summary>
        /// Parses a complete "Book Chapter:Verse" string into constituent parts.
        /// Handles numbered books like "1 Corinthians 13:4-7" and abbreviations.
        /// Returns null if the string cannot be parsed.
        /// </summary>
        public static (string BookName, string Reference)? ParseFullReference(Bible bible, string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;

            input = input.Trim();

            // Normalise dashes
            input = input.Replace('\u2013', '-').Replace('\u2014', '-');

            // Strategy: try progressively longer "book name" prefixes until
            // the remainder looks like a valid chapter:verse reference.
            // We split on spaces but must handle numbered-book prefixes like "1 " "2 " "3 " "I " "II " "III ".

            // Regex: capture the book part (letters, digits, spaces) then a chapter:verse tail.
            // The chapter:verse tail starts with a digit and contains :, -, , etc.
            // We use a greedy approach: find the last transition point from letters to the chapter digit.

            // First, handle trivial case: if the whole thing is a book name (whole chapter)
            var wholeBookResolved = ResolveBookName(bible, input);
            if (wholeBookResolved != null)
                return (wholeBookResolved, "");

            // Try to split "book" from "reference" by finding where the chapter number starts.
            // Walk from the end backwards to find the reference portion.
            // The reference portion is the last space-delimited token(s) that begin with a digit.
            string[] tokens = Regex.Split(input, @"\s+");

            // Try removing tokens from the end (those are the reference part)
            for (int refTokens = 1; refTokens < tokens.Length; refTokens++)
            {
                string refPart = string.Join(" ", tokens, tokens.Length - refTokens, refTokens);
                string bookPart = string.Join(" ", tokens, 0, tokens.Length - refTokens);

                // The reference part should start with a digit
                if (refPart.Length > 0 && char.IsDigit(refPart[0]))
                {
                    var resolved = ResolveBookName(bible, bookPart);
                    if (resolved != null)
                    {
                        return (resolved, refPart.Trim());
                    }
                }
            }

            return null;
        }
    }
}
