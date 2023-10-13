using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace VBA2CS
{
    public enum TokenType 
    {
        Identifier,
        Keyword,
        TypeKeyword,
        Operator,
        Literal,
        Comment,
        Separator,
        Whitespace,
        Unknown
    }

    public class Token
    {
        public TokenType Type { get; set; }
        public string Value { get; set; }
        public int Line { get; set; }
        public int Column { get; set; }
    }

    public class Tokenizer
    {
        private static readonly Regex TokenPattern = new Regex(@"('.*?$)|(\w+)|(\s+)|([=+\-*/(){};])|(-)|(\n)", RegexOptions.Multiline);
        private int _line;
        private int _column;

        public Tokenizer()
        {

        }

        public Token NextToken()
        {
            // 토큰 추출 로직
            // 예외 처리 및 에러 메시지 생성
            return null;
        }

        public List<Token> Tokenize(string code)
        {
            var tokens = new List<Token>();
            _line = 1;
            _column = 1;

            var matches = TokenPattern.Matches(code);
            bool isCommentLine = false;

            foreach (Match match in matches)
            {
                var token = new Token();

                if (isCommentLine)
                {
                    if (match.Groups[6].Success) // Newline
                    {
                        isCommentLine = false;
                    }
                }

                if (match.Groups[1].Success)
                {
                    isCommentLine = true;
                    token.Type = TokenType.Comment;
                    token.Value = match.Value;
                    token.Line = _line;
                    token.Column = _column;
                    tokens.Add(token);
                    continue;
                }
                else if (match.Groups[2].Success)
                {
                    if (IsKeyword(match.Groups[2].Value))
                    {
                        token.Type = TokenType.Keyword;
                    }
                    else
                    {
                        token.Type = TokenType.Identifier;
                    }
                }
                else if (match.Groups[3].Success)
                {
                    token.Type = TokenType.Whitespace;
                }
                else if (match.Groups[4].Success)
                {
                    token.Type = TokenType.Operator;
                }
                else if (match.Groups[5].Success)
                {
                    token.Type = TokenType.Operator;
                }
                else if (match.Groups[6].Success)
                {
                    _line++;
                }
                else
                {
                    token.Type = TokenType.Unknown;
                }

                token.Value = match.Value;
                token.Line = _line;
                token.Column = _column;

                tokens.Add(token);

                _column++;
            }

            return tokens;
        }

        private bool IsKeyword(string value)
        {
            return new HashSet<string> {
                "Sub", "Function", "Dim", "As", "If", "Then", "Else", "End",
                "For", "To", "Next", "While", "Do", "Loop", "Public", "Private",
                "Const", "Enum", "Type", "Case", "Select", "MsgBox", "Debug",
                "Print", "Integer", "String", "Boolean", "Option", "Explicit",
                "Each", "In", "Array", "ByVal", "ByRef", "True", "False", "Not",
                "And", "Or", "Mod", "Exit", "Continue", "GoTo", "Resume", "On",
                "Error", "Let", "Set", "Get", "Property", "Friend", "Static",
                "Global", "Call", "Me", "Nothing", "New", "Class", "Implements",
                "RaiseEvent", "Event", "Wend", "With", "ReDim", "Preserve",
                "Erase", "Step", "Until", "Is", "Like", "Optional", "ParamArray",
                "Variant", "Declare", "Lib", "Alias", "Long", "Single", "Double",
                "Currency", "Date", "Object", "Byte", "Boolean", "LongLong",
                "Integer", "WString", "String", "Decimal", "Stop", "End If",
                "ElseIf", "End Select", "End Sub", "End Function", "End Property",
                "End Type", "End Enum", "End Class", "End Interface", "End Module",
                "End With", "End Using", "End Namespace", "End Try", "End Structure"
            }.Contains(value);
        }
    }
}
