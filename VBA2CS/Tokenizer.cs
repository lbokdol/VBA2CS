using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace VBA2CS
{
    public class Token
    {
        public enum TokenType
        {
            Keyword,
            Operator,
            Literal,
            Identifier,
            Delimiter,
            Comment,
            Unknown
        }

        public TokenType Type { get; set; }
        public string Value { get; set; }
        public int LineNumber { get; set; }
        public int ColumnNumber { get; set; }

        public Token(TokenType type, string value, int lineNumber, int columnNumber)
        {
            Type = type;
            Value = value;
            LineNumber = lineNumber;
            ColumnNumber = columnNumber;
        }

        public override string ToString()
        {
            return $"[{Type}] {Value} (Line: {LineNumber}, Column: {ColumnNumber})";
        }
    }

    public class Tokenizer
    {
        private static readonly string[] Keywords = {
        "IF", "THEN", "ELSE", "END IF", "FOR", "NEXT", "DO", "LOOP", "WHILE", "WEND",
        "INTEGER", "LONG", "STRING", "DOUBLE", "VARIANT",
        "FUNCTION", "SUB", "END", "DIM", "SET", "LET", "PUBLIC", "PRIVATE"
    };

        private static readonly string[] Operators = {
        "+", "-", "*", "/", "^", "MOD",
        "=", "<", ">", "<=", ">=", "<>",
        "AND", "OR", "NOT", "XOR", "&"
    };

        private static readonly string[] Delimiters = {
        "(", ")", ",", ":", "."
    };

        public static List<Token> Tokenize(string code)
        {
            List<Token> tokens = new List<Token>();

            // Normalize the code (e.g., convert to uppercase for case-insensitivity)
            code = code.ToUpper();

            int lineNumber = 1;
            int columnNumber = 1;

            string[] lines = code.Split(new char[] { '\n' });
            string tmpLiteral = string.Empty;
            bool isLiteral = false;

            foreach (var line in lines)
            {
                string[] parts = line.Split(new char[] { ' ', '\t' }, StringSplitOptions.TrimEntries);

                foreach (var part in parts)
                {
                    if (part == "")
                        continue;

                    // Check for comments
                    if (part.StartsWith("'"))
                    {
                        tokens.Add(new Token(Token.TokenType.Comment, part, lineNumber, columnNumber));
                        break; // Skip the rest of the line
                    }

                    // Check for keywords
                    if (isLiteral)
                    {
                        tmpLiteral += part;
                        if (part.EndsWith("\""))
                        {
                            tokens.Add(new Token(Token.TokenType.Literal, tmpLiteral, lineNumber, columnNumber));
                            isLiteral = false;
                            tmpLiteral = string.Empty;
                        }
                    }
                    else if (Array.Exists(Keywords, keyword => keyword == part))
                    {
                        tokens.Add(new Token(Token.TokenType.Keyword, part, lineNumber, columnNumber));
                    }
                    // Check for operators
                    else if (Array.Exists(Operators, op => op == part))
                    {
                        tokens.Add(new Token(Token.TokenType.Operator, part, lineNumber, columnNumber));
                    }
                    // Check for delimiters
                    else if (Array.Exists(Delimiters, delimiter => delimiter == part))
                    {
                        tokens.Add(new Token(Token.TokenType.Delimiter, part, lineNumber, columnNumber));
                    }
                    // Check for string literals
                    else if (part.StartsWith("\"") && part.EndsWith("\"") && part.Length > 1)
                    {
                        tokens.Add(new Token(Token.TokenType.Literal, part, lineNumber, columnNumber));
                    }
                    else if (part.StartsWith("\""))
                    {
                        tmpLiteral = part;
                        isLiteral = true;
                    }
                    // Check for numeric literals
                    else if (Regex.IsMatch(part, @"^\d+(\.\d+)?$"))
                    {
                        tokens.Add(new Token(Token.TokenType.Literal, part, lineNumber, columnNumber));
                    }
                    // Check for date literals
                    else if (Regex.IsMatch(part, @"^#\d{1,2}/\d{1,2}/\d{4}#$"))
                    {
                        tokens.Add(new Token(Token.TokenType.Literal, part, lineNumber, columnNumber));
                    }
                    // Otherwise, treat as an identifier
                    else
                    {
                        tokens.Add(new Token(Token.TokenType.Identifier, part, lineNumber, columnNumber));
                    }

                    columnNumber += part.Length + 1; // +1 for the space or tab
                }

                lineNumber++;
                columnNumber = 1;
            }

            return tokens;
        }
    }
}
