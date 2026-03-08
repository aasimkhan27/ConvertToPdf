using System;
using System.Globalization;
using System.IO;
using System.Linq;

namespace SpreadsheetToPdf.Core
{
    internal sealed class CsvImportSettings
    {
        private CsvImportSettings(char delimiter, string decimalSeparator, string thousandSeparator)
        {
            Delimiter = delimiter;
            DecimalSeparator = decimalSeparator;
            ThousandSeparator = thousandSeparator;
        }

        public char Delimiter { get; }

        public string DecimalSeparator { get; }

        public string ThousandSeparator { get; }

        public bool UseOtherDelimiter => Delimiter != ',' && Delimiter != ';' && Delimiter != '\t';

        public static CsvImportSettings Detect(string inputPath)
        {
            // Sample up to first 25 non-empty lines for delimiter detection.
            string[] lines = File.ReadLines(inputPath)
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Take(25)
                .ToArray();

            char[] candidates = { ',', ';', '\t', '|', ':' };
            char selectedDelimiter = ',';
            int bestScore = int.MinValue;

            foreach (char candidate in candidates)
            {
                int[] counts = lines.Select(line => CountOccurrences(line, candidate)).ToArray();
                int nonZeroLines = counts.Count(c => c > 0);

                // Prefer delimiters that appear across many lines with stable counts.
                int score = nonZeroLines == 0 ? int.MinValue : (nonZeroLines * 100) - VariancePenalty(counts);
                if (score > bestScore)
                {
                    bestScore = score;
                    selectedDelimiter = candidate;
                }
            }

            // If nothing was detected confidently, keep comma default.
            if (bestScore == int.MinValue)
            {
                selectedDelimiter = ',';
            }

            CultureInfo culture = CultureInfo.CurrentCulture;
            string decimalSeparator = culture.NumberFormat.NumberDecimalSeparator;
            string thousandSeparator = culture.NumberFormat.NumberGroupSeparator;

            return new CsvImportSettings(selectedDelimiter, decimalSeparator, thousandSeparator);
        }

        private static int CountOccurrences(string input, char candidate)
        {
            int count = 0;
            bool insideQuotes = false;

            foreach (char value in input)
            {
                if (value == '"')
                {
                    insideQuotes = !insideQuotes;
                    continue;
                }

                if (!insideQuotes && value == candidate)
                {
                    count++;
                }
            }

            return count;
        }

        private static int VariancePenalty(int[] values)
        {
            if (values.Length == 0)
            {
                return 0;
            }

            double mean = values.Average();
            double variance = values.Sum(v => (v - mean) * (v - mean)) / values.Length;
            return (int)Math.Round(variance, MidpointRounding.AwayFromZero);
        }
    }
}
