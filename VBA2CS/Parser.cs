using System.Text;

namespace VBA2CS
{
    class Parser
    {
        private static readonly Dictionary<string, string> DataTypeMappings = new Dictionary<string, string>
        {
            { "INTEGER", "short" },
            { "LONG", "int" },
            { "STRING", "string" },
            { "DOUBLE", "double" },
            { "VARIANT", "object" },
            // ... 기타 데이터 타입 매핑 추가 ...
        };

        private static readonly Dictionary<string, string> OperatorMappings = new Dictionary<string, string>
        {
            { "&", "+" },
            { "AND", "&&" },
            { "OR", "||" },
            { "NOT", "!=" },
            { "XOR", "^" },
        };

        public static string ConvertTokensToCSharp(List<Token> tokens)
        {
            StringBuilder csharpCode = new StringBuilder();

            for (int i = 0; i < tokens.Count; i++)
            {
                Token token = tokens[i];

                switch (token.Type)
                {
                    case Token.TokenType.Keyword:
                        if (DataTypeMappings.ContainsKey(token.Value))
                        {
                            csharpCode.Append(DataTypeMappings[token.Value]);
                        }
                        else
                        {
                            switch (token.Value)
                            {
                                case "DIM":
                                    csharpCode.Append("var");
                                    break;
                                case "SUB":
                                    csharpCode.Append("void");
                                    break;
                                case "FUNCTION":
                                    csharpCode.Append("public");
                                    break;
                                case "IF":
                                    csharpCode.Append("if");
                                    break;
                                case "THEN":
                                    // 'THEN'은 C#에서 필요하지 않으므로 추가하지 않음
                                    break;
                                case "ELSE":
                                    csharpCode.Append("else");
                                    break;
                                case "END IF":
                                    // 'END IF'는 C#에서 '}'로 표현됨. 
                                    // 하지만, 이 로직에서는 단순 변환만 수행하므로 추가하지 않음
                                    break;
                                case "FOR":
                                    csharpCode.Append("for");
                                    break;
                                case "NEXT":
                                    // 'NEXT'는 C#에서 '}'로 표현됨.
                                    // 하지만, 이 로직에서는 단순 변환만 수행하므로 추가하지 않음
                                    break;
                                case "DO":
                                    csharpCode.Append("do");
                                    break;
                                case "LOOP":
                                    csharpCode.Append("while");
                                    break;
                                case "WHILE":
                                    csharpCode.Append("while");
                                    break;
                                case "WEND":
                                    // 'WEND'는 C#에서 '}'로 표현됨.
                                    // 하지만, 이 로직에서는 단순 변환만 수행하므로 추가하지 않음
                                    break;
                                case "SET":
                                case "LET":
                                    // VBA의 'SET'과 'LET'은 C#에서 필요하지 않으므로 추가하지 않음
                                    break;
                                case "PUBLIC":
                                    csharpCode.Append("public");
                                    break;
                                case "PRIVATE":
                                    csharpCode.Append("private");
                                    break;
                                case "END":
                                    // 'END'는 종료를 의미하는 키워드이므로, C#에서는 특별한 처리가 필요함
                                    // 여기서는 단순 변환만 수행하므로 추가하지 않음
                                    break;
                                default:
                                    csharpCode.Append(token.Value.ToLower()); // 기본적으로 VBA 키워드를 소문자로 변환
                                    break;
                            }
                        }
                        // ... 기타 키워드 변환 로직 추가 ...
                        break;

                    case Token.TokenType.Operator:
                        if (OperatorMappings.ContainsKey(token.Value))
                        {
                            csharpCode.Append(OperatorMappings[token.Value]);
                        }
                        else
                        {
                            csharpCode.Append(token.Value);
                        }
                        break;

                    case Token.TokenType.Literal:
                    case Token.TokenType.Identifier:
                    case Token.TokenType.Delimiter:
                        csharpCode.Append(token.Value);
                        break;

                    case Token.TokenType.Comment:
                        //csharpCode.Append("//" + token.Value.Substring(1)); // VBA의 '를 C#의 //로 변환
                        break;

                        // ... 기타 토큰 유형에 대한 변환 로직 추가 ...
                }

                csharpCode.Append(" "); // 각 토큰 사이에 공백 추가
            }

            return csharpCode.ToString();
        }
    }
}