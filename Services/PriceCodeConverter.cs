using System.Text;

namespace Excel2SQLExporter.Services;

/// <summary>
/// Converts a numeric price string into a cost code using a 10-character cipher key.
///
/// The key is stored in RegisteredUser.CostCode (e.g. "HOLYQURANF"), where each
/// character position maps to the digit at that index:
///   index 0 → 'H', 1 → 'O', 2 → 'L', 3 → 'Y', 4 → 'Q',
///         5 → 'U', 6 → 'R', 7 → 'A', 8 → 'N', 9 → 'F'
///
/// Algorithm (two passes):
///   Pass 1 — Substitute:  replace every digit with priceKey[digit]
///             Non-digits (e.g. '.') pass through unchanged.
///
///   Pass 2 — Compress:    for each run of the same character, the FIRST
///             occurrence is written as the character itself; every subsequent
///             consecutive occurrence is written as its 1-based run index.
///
/// Examples (priceKey = "HOLYQURANF"):
///   "50.00"   → substitute → "UH.HH"  → compress → "UH.H1"
///   "250.00"  → substitute → "LUH.HH" → compress → "LUH.H1"
///   "999.99"  → substitute → "FFF.FF" → compress → "F12.F1"
///   "100.00"  → substitute → "OHH.HH" → compress → "OH1.H1"
///   "1111.11" → substitute → "OOOO.OO"→ compress → "O123.O1"
/// </summary>
public static class PriceCodeConverter
{
    /// <summary>
    /// Converts a price string to a cost code using the supplied cipher key.
    /// </summary>
    /// <param name="price">
    ///   The price as a string, typically from <c>decimal.ToString("F2")</c>.
    ///   E.g. "250.00"
    /// </param>
    /// <param name="priceKey">
    ///   10-character cipher key from <c>RegisteredUser.CostCode</c>.
    ///   Must have at least 10 characters (indices 0-9 map to digits 0-9).
    /// </param>
    /// <returns>The encoded cost code string.</returns>
    public static string ConvertToCostCode(string price, string priceKey)
    {
        if (price.Length == 0)     return string.Empty;
        if (priceKey.Length < 10)  throw new ArgumentException("priceKey must have at least 10 characters.", nameof(priceKey));

        var sb = new StringBuilder(price.Length * 2);

        // ── Pass 1 + Pass 2 combined in a single loop ─────────────────────
        // Avoids allocating an intermediate char array.

        char prev = '\0';
        int  run  = 0;

        foreach (char c in price)
        {
            // Pass 1: substitute digit → cipher char; leave non-digits unchanged
            char mapped = char.IsDigit(c)
                ? priceKey[c - '0']     // e.g. '5' → priceKey[5], no string alloc
                : c;

            // Pass 2: consecutive-duplicate compression
            if (mapped == prev)
            {
                // Same char as previous — write run index (1-based for 2nd occurrence)
                sb.Append(++run);
            }
            else
            {
                // New char — write it and reset run counter
                sb.Append(mapped);
                prev = mapped;
                run  = 0;
            }
        }

        return sb.ToString();
    }
}
