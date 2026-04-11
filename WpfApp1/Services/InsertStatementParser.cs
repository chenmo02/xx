using System;
using System.Collections.Generic;

namespace WpfApp1.Services
{
    /// <summary>
    /// 解析 INSERT INTO 语句，提取字段列表和数据行。
    /// 支持多条 INSERT、批量 VALUES、字符串中的逗号/括号/关键字，以及 SQL 注释。
    /// </summary>
    public static class InsertStatementParser
    {
        public static (List<string> Headers, List<IReadOnlyList<string?>> Rows, string? Warning) Parse(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql))
                throw new InvalidOperationException("SQL 语句为空。");

            List<string>? headers = null;
            var rows = new List<IReadOnlyList<string?>>();
            string? warning = null;

            int index = 0;
            while (TryReadNextInsert(sql, ref index, out var currentHeaders, out var currentRows))
            {
                if (headers == null)
                {
                    headers = currentHeaders;
                }
                else if (!HeadersMatch(headers, currentHeaders) && warning == null)
                {
                    warning = "检测到多条 INSERT 的字段列表不一致，当前按第一条字段顺序解析。";
                }

                rows.AddRange(currentRows);
            }

            if (headers == null || headers.Count == 0)
                throw new InvalidOperationException("未找到有效的 INSERT INTO 语句（需要显式字段列表）。");

            if (rows.Count == 0)
                throw new InvalidOperationException("未解析到任何数据行，请检查 VALUES 部分格式。");

            return (headers, rows, warning);
        }

        private static bool TryReadNextInsert(string sql, ref int searchIndex, out List<string> headers, out List<IReadOnlyList<string?>> rows)
        {
            headers = [];
            rows = [];

            while (true)
            {
                int insertIndex = FindKeyword(sql, "INSERT", searchIndex);
                if (insertIndex < 0)
                    return false;

                int cursor = insertIndex + "INSERT".Length;
                if (!TryConsumeKeyword(sql, ref cursor, "INTO"))
                {
                    searchIndex = insertIndex + "INSERT".Length;
                    continue;
                }

                if (!TryFindNextChar(sql, ref cursor, '(', out int columnsStart))
                {
                    searchIndex = cursor;
                    continue;
                }

                cursor = columnsStart;
                string columnsText = ReadBalanced(sql, ref cursor, '(', ')');
                headers = ParseIdentifierList(columnsText);
                if (headers.Count == 0)
                {
                    searchIndex = cursor;
                    continue;
                }

                if (!TryConsumeKeyword(sql, ref cursor, "VALUES") &&
                    !TryConsumeKeyword(sql, ref cursor, "VALUE"))
                {
                    searchIndex = cursor;
                    continue;
                }

                rows = ReadValues(sql, ref cursor, headers.Count);
                searchIndex = cursor;
                return true;
            }
        }

        private static List<IReadOnlyList<string?>> ReadValues(string sql, ref int cursor, int expectedCount)
        {
            var rows = new List<IReadOnlyList<string?>>();

            while (cursor < sql.Length)
            {
                SkipTrivia(sql, ref cursor);
                while (cursor < sql.Length && sql[cursor] == ',')
                {
                    cursor++;
                    SkipTrivia(sql, ref cursor);
                }

                if (cursor >= sql.Length)
                    break;

                if (sql[cursor] == ';')
                {
                    cursor++;
                    break;
                }

                if (IsAtKeyword(sql, cursor, "INSERT"))
                    break;

                if (sql[cursor] != '(')
                {
                    cursor++;
                    continue;
                }

                string tuple = ReadBalanced(sql, ref cursor, '(', ')');
                rows.Add(ParseTuple(tuple, expectedCount));
            }

            return rows;
        }

        private static List<string> ParseIdentifierList(string value)
        {
            var result = new List<string>();
            foreach (var item in SplitTopLevel(value))
            {
                var name = item.Trim();
                if (string.IsNullOrEmpty(name))
                    continue;

                if (name.StartsWith("[", StringComparison.Ordinal) && name.EndsWith("]", StringComparison.Ordinal))
                    name = name[1..^1].Replace("]]", "]");
                else if (name.StartsWith("\"", StringComparison.Ordinal) && name.EndsWith("\"", StringComparison.Ordinal))
                    name = name[1..^1].Replace("\"\"", "\"");
                else if (name.StartsWith("`", StringComparison.Ordinal) && name.EndsWith("`", StringComparison.Ordinal))
                    name = name[1..^1].Replace("``", "`");

                if (!string.IsNullOrWhiteSpace(name))
                    result.Add(name);
            }

            return result;
        }

        private static List<string?> ParseTuple(string tuple, int expectedCount)
        {
            var result = new List<string?>();

            foreach (var item in SplitTopLevel(tuple))
            {
                var raw = item.Trim();
                if (raw.Equals("NULL", StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(null);
                    continue;
                }

                if (raw.Length >= 2 && raw[0] == '\'' && raw[^1] == '\'')
                {
                    result.Add(UnescapeSqlString(raw[1..^1]));
                    continue;
                }

                if (raw.Length >= 3 &&
                    (raw[0] == 'N' || raw[0] == 'n') &&
                    raw[1] == '\'' &&
                    raw[^1] == '\'')
                {
                    result.Add(UnescapeSqlString(raw[2..^1]));
                    continue;
                }

                result.Add(raw);
            }

            while (result.Count < expectedCount)
                result.Add(null);

            return result;
        }

        private static string UnescapeSqlString(string value)
        {
            return value.Replace("''", "'");
        }

        private static List<string> SplitTopLevel(string text)
        {
            var parts = new List<string>();
            int segmentStart = 0;
            int nestedParentheses = 0;
            bool inSingleQuote = false;
            bool inDoubleQuote = false;
            bool inBracketIdentifier = false;
            bool inBacktickIdentifier = false;

            for (int i = 0; i < text.Length; i++)
            {
                char current = text[i];

                if (inSingleQuote)
                {
                    if (current == '\'')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '\'')
                            i++;
                        else
                            inSingleQuote = false;
                    }

                    continue;
                }

                if (inDoubleQuote)
                {
                    if (current == '"')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '"')
                            i++;
                        else
                            inDoubleQuote = false;
                    }

                    continue;
                }

                if (inBracketIdentifier)
                {
                    if (current == ']')
                    {
                        if (i + 1 < text.Length && text[i + 1] == ']')
                            i++;
                        else
                            inBracketIdentifier = false;
                    }

                    continue;
                }

                if (inBacktickIdentifier)
                {
                    if (current == '`')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '`')
                            i++;
                        else
                            inBacktickIdentifier = false;
                    }

                    continue;
                }

                switch (current)
                {
                    case '\'':
                        inSingleQuote = true;
                        break;
                    case '"':
                        inDoubleQuote = true;
                        break;
                    case '[':
                        inBracketIdentifier = true;
                        break;
                    case '`':
                        inBacktickIdentifier = true;
                        break;
                    case '(':
                        nestedParentheses++;
                        break;
                    case ')':
                        if (nestedParentheses > 0)
                            nestedParentheses--;
                        break;
                    case ',':
                        if (nestedParentheses == 0)
                        {
                            parts.Add(text.Substring(segmentStart, i - segmentStart));
                            segmentStart = i + 1;
                        }
                        break;
                }
            }

            parts.Add(text.Substring(segmentStart));
            return parts;
        }

        private static string ReadBalanced(string text, ref int cursor, char openChar, char closeChar)
        {
            if (cursor >= text.Length || text[cursor] != openChar)
                throw new InvalidOperationException($"缺少 '{openChar}'。");

            int start = cursor + 1;
            int depth = 1;
            cursor++;

            bool inSingleQuote = false;
            bool inDoubleQuote = false;
            bool inBracketIdentifier = false;
            bool inBacktickIdentifier = false;

            while (cursor < text.Length)
            {
                char current = text[cursor];

                if (inSingleQuote)
                {
                    if (current == '\'')
                    {
                        if (cursor + 1 < text.Length && text[cursor + 1] == '\'')
                            cursor++;
                        else
                            inSingleQuote = false;
                    }

                    cursor++;
                    continue;
                }

                if (inDoubleQuote)
                {
                    if (current == '"')
                    {
                        if (cursor + 1 < text.Length && text[cursor + 1] == '"')
                            cursor++;
                        else
                            inDoubleQuote = false;
                    }

                    cursor++;
                    continue;
                }

                if (inBracketIdentifier)
                {
                    if (current == ']')
                    {
                        if (cursor + 1 < text.Length && text[cursor + 1] == ']')
                            cursor++;
                        else
                            inBracketIdentifier = false;
                    }

                    cursor++;
                    continue;
                }

                if (inBacktickIdentifier)
                {
                    if (current == '`')
                    {
                        if (cursor + 1 < text.Length && text[cursor + 1] == '`')
                            cursor++;
                        else
                            inBacktickIdentifier = false;
                    }

                    cursor++;
                    continue;
                }

                if (current == '\'')
                {
                    inSingleQuote = true;
                    cursor++;
                    continue;
                }

                if (current == '"')
                {
                    inDoubleQuote = true;
                    cursor++;
                    continue;
                }

                if (current == '[')
                {
                    inBracketIdentifier = true;
                    cursor++;
                    continue;
                }

                if (current == '`')
                {
                    inBacktickIdentifier = true;
                    cursor++;
                    continue;
                }

                if (current == openChar)
                    depth++;
                else if (current == closeChar)
                    depth--;

                cursor++;
                if (depth == 0)
                    return text.Substring(start, cursor - start - 1);
            }

            throw new InvalidOperationException($"未找到匹配的 '{closeChar}'。");
        }

        private static int FindKeyword(string text, string keyword, int startIndex)
        {
            int index = startIndex;
            while (index < text.Length)
            {
                SkipTrivia(text, ref index);
                if (index >= text.Length)
                    break;

                char current = text[index];
                if (current == '\'')
                {
                    SkipSingleQuotedString(text, ref index);
                    continue;
                }

                if (current == '"')
                {
                    SkipDoubleQuotedText(text, ref index);
                    continue;
                }

                if (current == '[')
                {
                    SkipBracketIdentifier(text, ref index);
                    continue;
                }

                if (current == '`')
                {
                    SkipBacktickIdentifier(text, ref index);
                    continue;
                }

                if (IsAtKeyword(text, index, keyword))
                    return index;

                index++;
            }

            return -1;
        }

        private static bool TryConsumeKeyword(string text, ref int cursor, string keyword)
        {
            SkipTrivia(text, ref cursor);
            if (!IsAtKeyword(text, cursor, keyword))
                return false;

            cursor += keyword.Length;
            return true;
        }

        private static bool TryFindNextChar(string text, ref int cursor, char target, out int position)
        {
            while (cursor < text.Length)
            {
                SkipTrivia(text, ref cursor);
                if (cursor >= text.Length)
                    break;

                char current = text[cursor];
                if (current == '\'')
                {
                    SkipSingleQuotedString(text, ref cursor);
                    continue;
                }

                if (current == '"')
                {
                    SkipDoubleQuotedText(text, ref cursor);
                    continue;
                }

                if (current == '[')
                {
                    SkipBracketIdentifier(text, ref cursor);
                    continue;
                }

                if (current == '`')
                {
                    SkipBacktickIdentifier(text, ref cursor);
                    continue;
                }

                if (current == target)
                {
                    position = cursor;
                    return true;
                }

                cursor++;
            }

            position = -1;
            return false;
        }

        private static void SkipTrivia(string text, ref int cursor)
        {
            while (cursor < text.Length)
            {
                if (char.IsWhiteSpace(text[cursor]))
                {
                    cursor++;
                    continue;
                }

                if (cursor + 1 < text.Length && text[cursor] == '-' && text[cursor + 1] == '-')
                {
                    cursor += 2;
                    while (cursor < text.Length && text[cursor] != '\r' && text[cursor] != '\n')
                        cursor++;
                    continue;
                }

                if (cursor + 1 < text.Length && text[cursor] == '/' && text[cursor + 1] == '*')
                {
                    cursor += 2;
                    while (cursor + 1 < text.Length && !(text[cursor] == '*' && text[cursor + 1] == '/'))
                        cursor++;

                    if (cursor + 1 < text.Length)
                        cursor += 2;

                    continue;
                }

                break;
            }
        }

        private static void SkipSingleQuotedString(string text, ref int cursor)
        {
            cursor++;
            while (cursor < text.Length)
            {
                if (text[cursor] == '\'')
                {
                    if (cursor + 1 < text.Length && text[cursor + 1] == '\'')
                        cursor += 2;
                    else
                    {
                        cursor++;
                        break;
                    }
                }
                else
                {
                    cursor++;
                }
            }
        }

        private static void SkipDoubleQuotedText(string text, ref int cursor)
        {
            cursor++;
            while (cursor < text.Length)
            {
                if (text[cursor] == '"')
                {
                    if (cursor + 1 < text.Length && text[cursor + 1] == '"')
                        cursor += 2;
                    else
                    {
                        cursor++;
                        break;
                    }
                }
                else
                {
                    cursor++;
                }
            }
        }

        private static void SkipBracketIdentifier(string text, ref int cursor)
        {
            cursor++;
            while (cursor < text.Length)
            {
                if (text[cursor] == ']')
                {
                    if (cursor + 1 < text.Length && text[cursor + 1] == ']')
                        cursor += 2;
                    else
                    {
                        cursor++;
                        break;
                    }
                }
                else
                {
                    cursor++;
                }
            }
        }

        private static void SkipBacktickIdentifier(string text, ref int cursor)
        {
            cursor++;
            while (cursor < text.Length)
            {
                if (text[cursor] == '`')
                {
                    if (cursor + 1 < text.Length && text[cursor + 1] == '`')
                        cursor += 2;
                    else
                    {
                        cursor++;
                        break;
                    }
                }
                else
                {
                    cursor++;
                }
            }
        }

        private static bool IsAtKeyword(string text, int index, string keyword)
        {
            if (index < 0 || index + keyword.Length > text.Length)
                return false;

            if (!text.AsSpan(index, keyword.Length).Equals(keyword, StringComparison.OrdinalIgnoreCase))
                return false;

            bool leftBoundary = index == 0 || !IsIdentifierChar(text[index - 1]);
            bool rightBoundary = index + keyword.Length >= text.Length || !IsIdentifierChar(text[index + keyword.Length]);
            return leftBoundary && rightBoundary;
        }

        private static bool IsIdentifierChar(char value)
        {
            return char.IsLetterOrDigit(value) || value == '_';
        }

        private static bool HeadersMatch(List<string> left, List<string> right)
        {
            if (left.Count != right.Count)
                return false;

            for (int i = 0; i < left.Count; i++)
            {
                if (!string.Equals(left[i], right[i], StringComparison.OrdinalIgnoreCase))
                    return false;
            }

            return true;
        }
    }
}
