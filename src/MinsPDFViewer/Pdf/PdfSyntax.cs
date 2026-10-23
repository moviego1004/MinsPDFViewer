using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace MinsPDFViewer
{
    /// <summary>
    /// 증분 저장용 최소 PDF 구문 도구. 사전(<c>&lt;&lt; &gt;&gt;</c>)의 최상위 키만 찾고 바꾼다.
    /// 중첩 사전·배열·문자열(괄호, 16진)·이름·간접 참조를 건너뛸 줄 알아서 정규식처럼 중첩에서 깨지지 않는다.
    /// </summary>
    internal static class PdfSyntax
    {
        private static readonly Regex RefRegex = new(@"^\s*(\d+)\s+(\d+)\s+R\s*$", RegexOptions.Compiled);

        public static bool IsWhite(char c) => c is ' ' or '\t' or '\r' or '\n' or '\f' or '\0';

        public static bool IsDelimiter(char c) => c is '(' or ')' or '<' or '>' or '[' or ']' or '{' or '}' or '/' or '%';

        public static int SkipWhite(string s, int i)
        {
            while (i < s.Length)
            {
                char c = s[i];
                if (IsWhite(c))
                {
                    i++;
                    continue;
                }

                if (c == '%')
                {
                    while (i < s.Length && s[i] != '\r' && s[i] != '\n')
                        i++;
                    continue;
                }

                break;
            }

            return i;
        }

        /// <summary>i 위치(공백 허용)에서 시작하는 객체 하나를 건너뛰고 그 다음 위치를 돌려준다.</summary>
        public static int SkipObject(string s, int i)
        {
            i = SkipWhite(s, i);
            if (i >= s.Length)
                return i;

            char c = s[i];
            if (c == '<' && i + 1 < s.Length && s[i + 1] == '<')
            {
                i += 2;
                while (true)
                {
                    i = SkipWhite(s, i);
                    if (i >= s.Length)
                        throw new FormatException("Unterminated PDF dictionary.");
                    if (s[i] == '>' && i + 1 < s.Length && s[i + 1] == '>')
                        return i + 2;
                    i = SkipObject(s, i);
                }
            }

            if (c == '[')
            {
                i++;
                while (true)
                {
                    i = SkipWhite(s, i);
                    if (i >= s.Length)
                        throw new FormatException("Unterminated PDF array.");
                    if (s[i] == ']')
                        return i + 1;
                    i = SkipObject(s, i);
                }
            }

            if (c == '(')
            {
                int depth = 0;
                for (; i < s.Length; i++)
                {
                    char ch = s[i];
                    if (ch == '\\')
                    {
                        i++;
                        continue;
                    }

                    if (ch == '(')
                    {
                        depth++;
                    }
                    else if (ch == ')')
                    {
                        depth--;
                        if (depth == 0)
                            return i + 1;
                    }
                }

                throw new FormatException("Unterminated PDF literal string.");
            }

            if (c == '<')
            {
                int end = s.IndexOf('>', i + 1);
                if (end < 0)
                    throw new FormatException("Unterminated PDF hex string.");
                return end + 1;
            }

            if (c == '/')
            {
                i++;
                return ReadRegular(s, i);
            }

            if (c is ')' or '>' or ']' or '{' or '}')
                return i + 1;

            // 숫자, 키워드(true/false/null) 또는 간접 참조 "N G R"
            int tokenEnd = ReadRegular(s, i);
            if (tokenEnd == i)
                return i + 1;

            if (IsUnsignedInteger(s, i, tokenEnd))
            {
                int j = SkipWhite(s, tokenEnd);
                int genEnd = ReadRegular(s, j);
                if (genEnd > j && IsUnsignedInteger(s, j, genEnd))
                {
                    int k = SkipWhite(s, genEnd);
                    if (k < s.Length && s[k] == 'R' && (k + 1 >= s.Length || IsWhite(s[k + 1]) || IsDelimiter(s[k + 1])))
                        return k + 1;
                }
            }

            return tokenEnd;
        }

        private static int ReadRegular(string s, int i)
        {
            while (i < s.Length && !IsWhite(s[i]) && !IsDelimiter(s[i]))
                i++;
            return i;
        }

        private static bool IsUnsignedInteger(string s, int start, int end)
        {
            if (end <= start)
                return false;
            for (int k = start; k < end; k++)
            {
                if (s[k] < '0' || s[k] > '9')
                    return false;
            }
            return true;
        }

        /// <summary>객체 본문(사전, 또는 사전 + stream)에서 맨 앞 사전만 잘라 낸다.</summary>
        public static string ExtractLeadingDictionary(string content)
        {
            int start = content.IndexOf("<<", StringComparison.Ordinal);
            if (start < 0)
                throw new FormatException("PDF dictionary was not found.");
            int end = SkipObject(content, start);
            return content.Substring(start, end - start);
        }

        public static bool TryFindEntry(string dict, string key, out int keyStart, out int valueStart, out int valueEnd)
        {
            keyStart = valueStart = valueEnd = -1;
            int i = SkipWhite(dict, 0);
            if (i + 1 >= dict.Length || dict[i] != '<' || dict[i + 1] != '<')
                return false;
            i += 2;

            while (true)
            {
                i = SkipWhite(dict, i);
                if (i >= dict.Length)
                    return false;
                if (dict[i] == '>' && i + 1 < dict.Length && dict[i + 1] == '>')
                    return false;

                if (dict[i] != '/')
                {
                    // 키 자리에 이름이 아닌 것이 있으면 건너뛴다 (깨진 사전에 대한 관용)
                    i = SkipObject(dict, i);
                    continue;
                }

                int ks = i;
                int ke = SkipObject(dict, i);
                int vs = SkipWhite(dict, ke);
                int ve = SkipObject(dict, vs);
                if (ke - ks - 1 == key.Length && string.CompareOrdinal(dict, ks + 1, key, 0, key.Length) == 0)
                {
                    keyStart = ks;
                    valueStart = vs;
                    valueEnd = ve;
                    return true;
                }

                i = ve;
            }
        }

        public static string? GetValue(string dict, string key)
        {
            return TryFindEntry(dict, key, out _, out int vs, out int ve)
                ? dict.Substring(vs, ve - vs).Trim()
                : null;
        }

        /// <summary>최상위 키 값을 바꾸거나, 없으면 사전 끝에 추가한다.</summary>
        public static string SetValue(string dict, string key, string value)
        {
            if (TryFindEntry(dict, key, out int ks, out _, out int ve))
                return dict.Substring(0, ks) + "/" + key + " " + value + dict.Substring(ve);

            int start = SkipWhite(dict, 0);
            int end = SkipObject(dict, start); // ">>" 바로 다음
            int insertAt = end - 2;
            return dict.Substring(0, insertAt) + " /" + key + " " + value + " " + dict.Substring(insertAt);
        }

        public static bool TryParseReference(string? value, out int number, out int generation)
        {
            number = generation = 0;
            if (value == null)
                return false;
            var m = RefRegex.Match(value);
            if (!m.Success)
                return false;
            number = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            generation = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            return true;
        }

        /// <summary>배열 문자열 "[a b c]"의 원소들을 문자열로 돌려준다 (간접 참조는 한 원소).</summary>
        public static List<string> ParseArrayItems(string array)
        {
            var items = new List<string>();
            int i = SkipWhite(array, 0);
            if (i >= array.Length || array[i] != '[')
                return items;
            i++;
            while (true)
            {
                i = SkipWhite(array, i);
                if (i >= array.Length || array[i] == ']')
                    break;
                int end = SkipObject(array, i);
                items.Add(array.Substring(i, end - i));
                i = end;
            }
            return items;
        }

        public static bool TryParseNumber(string token, out double value)
        {
            return double.TryParse(token.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }
    }
}
