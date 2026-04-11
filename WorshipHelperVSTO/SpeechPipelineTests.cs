// ============================================================================
// SpeechPipelineTests.cs
//
// Console-based test harness for the speech-to-scripture pipeline.
// Run this as a standalone console app to verify the number converter,
// Bible reference detector, and the full pipeline WITHOUT needing PowerPoint.
//
// Usage:
//   1. Create a new C# Console App (.NET Framework 4.7.2)
//   2. Add the four .cs files (SpokenNumberConverter, BibleReferenceDetector,
//      SpeechListener, SpeechToScriptureService)
//   3. Add this test file
//   4. Add reference to System.Speech
//   5. Add log4net NuGet package (or stub out the ILog calls)
//   6. Run — it will execute all tests, then optionally do live mic testing.
//
// ============================================================================

using System;
using System.Collections.Generic;

namespace WorshipHelperVSTO
{
    class SpeechPipelineTests
    {
        static int passed = 0;
        static int failed = 0;

        static void Main(string[] args)
        {
            Console.WriteLine("═══════════════════════════════════════════════════════════");
            Console.WriteLine("  WorshipHelper Speech-to-Scripture Pipeline — Test Suite  ");
            Console.WriteLine("═══════════════════════════════════════════════════════════");
            Console.WriteLine();

            TestSpokenNumberConverter();
            TestSpokenToReferenceFragment();
            TestBibleReferenceDetector();
            TestEdgeCases();

            Console.WriteLine();
            Console.WriteLine("═══════════════════════════════════════════════════════════");
            Console.ForegroundColor = failed == 0 ? ConsoleColor.Green : ConsoleColor.Red;
            Console.WriteLine($"  Results: {passed} passed, {failed} failed");
            Console.ResetColor();
            Console.WriteLine("═══════════════════════════════════════════════════════════");

            // Interactive mode
            Console.WriteLine();
            Console.WriteLine("Enter spoken phrases to test detection (or 'quit' to exit):");
            Console.WriteLine("Examples: 'turn to john three sixteen'");
            Console.WriteLine("          'first corinthians thirteen verse four'");
            Console.WriteLine();

            while (true)
            {
                Console.Write("> ");
                string input = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(input) || input.Trim().ToLower() == "quit")
                    break;

                var results = BibleReferenceDetector.Detect(input);
                if (results.Count == 0)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("  No Bible reference detected.");
                    Console.ResetColor();
                }
                else
                {
                    foreach (var r in results)
                    {
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine($"  → {r.NormalisedReference}  (confidence: {r.Confidence:F2}, matched: \"{r.MatchedRawText}\")");
                        Console.ResetColor();
                    }
                }
            }
        }

        // ===================================================================
        // SpokenNumberConverter tests
        // ===================================================================

        static void TestSpokenNumberConverter()
        {
            Console.WriteLine("── SpokenNumberConverter.WordsToNumber ──");

            AssertNumber("three", 3);
            AssertNumber("sixteen", 16);
            AssertNumber("twenty one", 21);
            AssertNumber("twenty", 20);
            AssertNumber("one hundred", 100);
            AssertNumber("one hundred and three", 103);
            AssertNumber("one hundred twenty one", 121);
            AssertNumber("one hundred and nineteen", 119);
            AssertNumber("one oh five", 105);
            AssertNumber("one oh three", 103);
            AssertNumber("fifty", 50);
            AssertNumber("ninety nine", 99);
            AssertNumber("twelve", 12);
            AssertNumber("zero", 0);
            AssertNumber("one", 1);

            Console.WriteLine();
            Console.WriteLine("── SpokenNumberConverter.OrdinalToDigit ──");

            AssertEqual(SpokenNumberConverter.OrdinalToDigit("first"), "1", "first → 1");
            AssertEqual(SpokenNumberConverter.OrdinalToDigit("second"), "2", "second → 2");
            AssertEqual(SpokenNumberConverter.OrdinalToDigit("third"), "3", "third → 3");
            AssertEqual(SpokenNumberConverter.OrdinalToDigit("1st"), "1", "1st → 1");
            AssertEqual(SpokenNumberConverter.OrdinalToDigit("2nd"), "2", "2nd → 2");
            AssertEqual(SpokenNumberConverter.OrdinalToDigit("3rd"), "3", "3rd → 3");

            Console.WriteLine();
        }

        // ===================================================================
        // SpokenToReferenceFragment tests
        // ===================================================================

        static void TestSpokenToReferenceFragment()
        {
            Console.WriteLine("── SpokenNumberConverter.SpokenToReferenceFragment ──");

            AssertFragment("three sixteen", "3:16");
            AssertFragment("twenty three", "23");
            AssertFragment("chapter three verse sixteen", "3:16");
            AssertFragment("fifty three five", "53:5");
            AssertFragment("one hundred and nineteen verse one hundred and five", "119:105");
            AssertFragment("three sixteen to eighteen", "3:16-18");
            AssertFragment("three sixteen and seventeen", "3:16-17");
            AssertFragment("thirteen verse four", "13:4");
            AssertFragment("three", "3");

            Console.WriteLine();
        }

        // ===================================================================
        // BibleReferenceDetector tests
        // ===================================================================

        static void TestBibleReferenceDetector()
        {
            Console.WriteLine("── BibleReferenceDetector.Detect ──");

            // Core examples from the requirements
            AssertDetection("john three sixteen", "John 3:16");
            AssertDetection("first corinthians thirteen verse four", "1 Corinthians 13:4");
            AssertDetection("psalm twenty three", "Psalms 23");
            AssertDetection("isaiah fifty three five", "Isaiah 53:5");
            AssertDetection("john three sixteen to eighteen", "John 3:16-18");

            // With preamble filler words
            AssertDetection("read john three sixteen", "John 3:16");
            AssertDetection("turn to first corinthians thirteen verse four", "1 Corinthians 13:4");
            AssertDetection("let's go to psalm twenty three", "Psalms 23");
            AssertDetection("let's turn to isaiah fifty three five", "Isaiah 53:5");
            AssertDetection("please open to revelation twenty one three", "Revelation 21:3");

            // Numbered book variants
            AssertDetection("second corinthians five seventeen", "2 Corinthians 5:17");
            AssertDetection("1st corinthians thirteen four", "1 Corinthians 13:4");
            AssertDetection("third john one four", "3 John 1:4");
            AssertDetection("second timothy three sixteen", "2 Timothy 3:16");

            // Various books
            AssertDetection("genesis one one", "Genesis 1:1");
            AssertDetection("romans eight twenty eight", "Romans 8:28");
            AssertDetection("ephesians two eight", "Ephesians 2:8");
            AssertDetection("philippians four thirteen", "Philippians 4:13");
            AssertDetection("hebrews eleven one", "Hebrews 11:1");
            AssertDetection("matthew five three", "Matthew 5:3");
            AssertDetection("proverbs three five", "Proverbs 3:5");

            Console.WriteLine();
        }

        // ===================================================================
        // Edge cases
        // ===================================================================

        static void TestEdgeCases()
        {
            Console.WriteLine("── Edge Cases ──");

            // Should NOT detect anything in generic speech
            AssertNoDetection("the weather is nice today");
            AssertNoDetection("please pass the salt");
            AssertNoDetection("welcome to church everyone");
            AssertNoDetection("let us pray");
            AssertNoDetection("good morning");

            // Partial book names that shouldn't trigger
            AssertNoDetection("the mark of the beast");  // "mark" alone + no numbers = no trigger

            Console.WriteLine();
        }

        // ===================================================================
        // Assertion helpers
        // ===================================================================

        static void AssertNumber(string input, int expected)
        {
            int? result = SpokenNumberConverter.WordsToNumber(input);
            if (result.HasValue && result.Value == expected)
            {
                passed++;
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"  ✓ \"{input}\" → {result.Value}");
            }
            else
            {
                failed++;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"  ✗ \"{input}\" → {result?.ToString() ?? "null"} (expected {expected})");
            }
            Console.ResetColor();
        }

        static void AssertFragment(string input, string expected)
        {
            string result = SpokenNumberConverter.SpokenToReferenceFragment(input);
            if (result == expected)
            {
                passed++;
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"  ✓ \"{input}\" → \"{result}\"");
            }
            else
            {
                failed++;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"  ✗ \"{input}\" → \"{result ?? "null"}\" (expected \"{expected}\")");
            }
            Console.ResetColor();
        }

        static void AssertDetection(string input, string expectedRef)
        {
            var best = BibleReferenceDetector.DetectBest(input);
            if (best != null && best.NormalisedReference == expectedRef)
            {
                passed++;
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"  ✓ \"{input}\" → \"{best.NormalisedReference}\" (conf={best.Confidence:F2})");
            }
            else
            {
                failed++;
                Console.ForegroundColor = ConsoleColor.Red;
                string actual = best?.NormalisedReference ?? "null";
                Console.WriteLine($"  ✗ \"{input}\" → \"{actual}\" (expected \"{expectedRef}\")");
            }
            Console.ResetColor();
        }

        static void AssertNoDetection(string input)
        {
            var best = BibleReferenceDetector.DetectBest(input);
            if (best == null)
            {
                passed++;
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"  ✓ \"{input}\" → (no detection) ✓");
            }
            else
            {
                failed++;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"  ✗ \"{input}\" → \"{best.NormalisedReference}\" (expected no detection)");
            }
            Console.ResetColor();
        }

        static void AssertEqual(string actual, string expected, string label)
        {
            if (actual == expected)
            {
                passed++;
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"  ✓ {label}");
            }
            else
            {
                failed++;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"  ✗ {label}: got \"{actual ?? "null"}\" (expected \"{expected}\")");
            }
            Console.ResetColor();
        }
    }
}
