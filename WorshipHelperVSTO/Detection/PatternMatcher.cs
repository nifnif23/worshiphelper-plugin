// ============================================================================
// Detection/PatternMatcher.cs
// Facade over the existing BibleReferenceDetector so new Pipeline code can
// inject a clean IPatternMatcher interface, while the 500-line regex engine
// stays exactly where it is.
// ============================================================================
using System.Collections.Generic;

namespace WorshipHelperVSTO.Detection
{
    public interface IPatternMatcher
    {
        IReadOnlyList<DetectedReference> Detect(string text);
        DetectedReference DetectBest(string text);
    }

    public sealed class PatternMatcher : IPatternMatcher
    {
        public IReadOnlyList<DetectedReference> Detect(string text) =>
            BibleReferenceDetector.Detect(text);

        public DetectedReference DetectBest(string text) =>
            BibleReferenceDetector.DetectBest(text);
    }
}
