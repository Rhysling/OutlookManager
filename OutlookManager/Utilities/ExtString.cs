using System;
using System.Diagnostics.CodeAnalysis;

namespace OutlookMailParser.Utilities
{
	public static class ExtString
	{
		public static string Left(this string value, int length)
		{
			if (value == null) return "";
			int len = value.Length;
			if (length == 0) return "";
			if (len <= length) return value;
			return value[..length];
		}

		public static string Right(this string value, int length)
		{
			if (value == null) return "";
			int len = value.Length;
			if (length == 0) return "";
			if (len <= length) return value;
			return value[^length..];
		}

		public static string LeadStr(this string strIn, int ShowLen)
		{
			//'*** Shows the first n letters of a string,
			//'    followed by "..." if more than n
			int ls = strIn.Length;
			if (ls <= ShowLen) return strIn;

			return string.Concat(strIn.AsSpan(0, ShowLen), "...");
		}

		// Returns null when given an empty string.
		public static string? NullWhenEmpty(this string? strIn)
		{
			return String.IsNullOrWhiteSpace(strIn) ? null : strIn.Trim();
		}


		public static bool IsEmpty([NotNullWhen(false)] this string? val)
		{
			return String.IsNullOrWhiteSpace(val);
		}

		public static bool IsNotEmpty([NotNullWhen(true)] this string? val)
		{
			return !String.IsNullOrWhiteSpace(val);
		}


		public static bool StrictEqual(this string? val, string? other)
		{
			return String.CompareOrdinal(val, other) == 0;
		}

		public static bool LooseEqual(this string? val, string? other)
		{
			return String.Compare(val, other, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.CompareOptions.IgnoreCase) == 0;
		}

		public static bool IsGreaterThanOrEqual(this string? val, string? other)
		{
			return String.CompareOrdinal(val, other) >= 0;
		}

		public static bool IsLessThanOrEqual(this string val, string other)
		{
			return String.CompareOrdinal(val, other) <= 0;
		}

		public static bool StartsWithsCaseInsensitive(this string val, string other)
		{
			if (String.IsNullOrEmpty(val) || String.IsNullOrEmpty(other)) return false;

			return (val.IndexOf(other, StringComparison.CurrentCultureIgnoreCase) == 0);
		}

		public static bool ContainsCaseInsensitive(this string val, string other)
		{
			if (String.IsNullOrEmpty(val) || String.IsNullOrEmpty(other)) return false;

			return (val.IndexOf(other, StringComparison.CurrentCultureIgnoreCase) > -1);
		}


		public static bool RegExIsMatch(this string val, string pattern)
		{
			if (val == null || pattern == null)
				return false;

			return System.Text.RegularExpressions.Regex.IsMatch(val, pattern);
		}

		public static string RegExMatchValue(this string val, string pattern)
		{
			if (val == null || pattern == null)
				return "";

			return System.Text.RegularExpressions.Regex.Match(val, pattern).Value;
		}


		public static DateTime GmtFromTwitter(this string val)
		{
			//Fri Mar 18 21:19:23 +0000 2011
			//0123456789012345678901234567890

			return DateTime.Parse(val[..3] + ", " + val.Substring(8, 2) + " " + val.Substring(4, 3) + " " + val.Substring(26, 4) + " " + val.Substring(11, 8) + " GMT", null, System.Globalization.DateTimeStyles.RoundtripKind);

			//"Mon, 15 Sep 2008 09:30:41 GMT"

		}

	}
}
