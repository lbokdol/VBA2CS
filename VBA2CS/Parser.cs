using System.Text;
using System.Xml.Linq;
using static VBA2CS.Token;

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
            { "<>", "!=" },
        };

        private int _position = 0;
        private List<Token> _tokens;

        public Parser(List<Token> tokens)
        {
            _tokens = tokens;
        }

        public ASTNode Parse()
        {
            return ParseStatement();
        }

        private ASTNode ParseStatement()
        {
            var token = LookAhead();
            Console.WriteLine($"token : {token.Value} !!!{_position}");

            if (token.Type == Token.TokenType.Keyword && token.Value == "FUNCTION")
            {
                return ParseFunction();
            }
            else if (token.Type == Token.TokenType.Keyword && token.Value == "SUB")
            {
                return ParseSubroutine();
            }
            else if (token.Type == Token.TokenType.Keyword && token.Value == "DIM")
            {
                return ParseVariableDeclaration();
            }
            else if (token.Type == Token.TokenType.Keyword && token.Value == "IF")
            {
                return ParseIfStatement();
            }
            else if (token.Type == Token.TokenType.Operator && token.Value == "=")
            {
                return ParseAssignment();
            }
            else if (token.Type == Token.TokenType.Comment)
            {
                _position++;
                return new CommentNode() { Value = token.Value, };
            }
            else if (token.Type == Token.TokenType.Identifier)
            {
                _position++;
                return new CommentNode() { Value = "", };
            }
            // ... 기타 문장 타입 파싱 ...

            return null; // 오류 처리
        }

        private FunctionNode ParseFunction()
        {
            Match(Token.TokenType.Keyword, "FUNCTION");
            var functionName = Match(Token.TokenType.Identifier).Value;
            var functionNode = new FunctionNode { Name = functionName };

            Match(Token.TokenType.Delimiter, "(");

            while (LookAhead().Type != Token.TokenType.Delimiter && LookAhead().Value != ")")
            {
                var parameter = ParseVariable();
                if (parameter == null)
                    break;

                functionNode.Parameters.Add(parameter);
                if (LookAhead().Value == ",")
                {
                    Match(Token.TokenType.Delimiter, ",");
                }
            }

            Match(Token.TokenType.Delimiter, ")");

            var As = Match(Token.TokenType.Keyword, "AS");
            if (As != null && LookAhead().Type == TokenType.Keyword)
                functionNode.Type = LookAhead().Value;

            while (_position < _tokens.Count && _tokens[_position].Value.Contains("END") == false)
            {
                var node = ParseStatement();
                if (node == null)
                {
                    _position++;
                    continue;
                }
                    
                functionNode.Body.Add(node);
                
            }

            return functionNode;
        }

        private SubroutineNode ParseSubroutine() 
        {
            Match(Token.TokenType.Keyword, "SUB");
            var subroutineName = Match(Token.TokenType.Identifier).Value;

            var subroutineNode = new SubroutineNode { Name = subroutineName };

            Match(Token.TokenType.Delimiter, "(");

            while (LookAhead().Type != Token.TokenType.Delimiter && LookAhead().Value != ")")
            {
                var parameter = ParseVariable();
                if (parameter == null)
                    break;

                subroutineNode.Parameters.Add(parameter);
                if (LookAhead().Value == ",")
                {
                    Match(Token.TokenType.Delimiter, ",");
                }
            }

            Match(Token.TokenType.Delimiter, ")");

            while (_position < _tokens.Count && _tokens[_position].Value.Contains("END") == false)
            {
                subroutineNode.Body.Add(ParseStatement());
            }

            return subroutineNode;
        }

        private IfNode ParseIfStatement()
        {
            // IF 키워드와 조건식 파싱
            Match(Token.TokenType.Keyword, "IF");
            var condition = ParseExpression(); // 조건식 파싱 로직을 호출
            Match(Token.TokenType.Keyword, "THEN");

            var ifNode = new IfNode { Condition = condition };

            // THEN 블록 내의 문장들 파싱
            while (LookAhead().Value != "ELSE" && LookAhead().Value != "END")
            {
                ifNode.TrueBranch.Add(ParseStatement());
            }

            // ELSE 블록이 있는 경우
            if (LookAhead().Value == "ELSE")
            {
                Match(Token.TokenType.Keyword, "ELSE");

                while (LookAhead().Value != "END")
                {
                    ifNode.FalseBranch.Add(ParseStatement());
                }
            }

            // END IF 키워드 파싱
            Match(Token.TokenType.Keyword, "END");
            _position++;

            return ifNode;
        }

        private VariableNode ParseVariableDeclaration()
        {
            // DIM 키워드를 확인하고 넘어갑니다.
            Match(Token.TokenType.Keyword, "DIM");

            // 변수 이름을 가져옵니다.
            var variableNameToken = Match(Token.TokenType.Identifier);
            string variableName = variableNameToken.Value;

            // AS 키워드를 확인하고 넘어갑니다.
            Match(Token.TokenType.Keyword, "AS");

            // 변수 타입을 가져옵니다.
            var variableTypeToken = Match(Token.TokenType.Keyword); // 또는 Identifier, 상황에 따라 다름
            string variableType = variableTypeToken.Value;

            // AST 노드를 생성하고 채웁니다.
            var variableDeclarationNode = new VariableNode
            {
                Name = variableName,
                Type = variableType
            };

            return variableDeclarationNode;
        }

        private AssignmentNode ParseAssignment()
        {
            // 변수 이름을 가져옵니다.
            var variableNameToken = Match(Token.TokenType.Identifier, beforePosition:1);
            string variableName = variableNameToken.Value;

            // 할당 연산자를 확인하고 넘어갑니다.
            Match(Token.TokenType.Operator, "=");

            // 할당되는 값을 파싱합니다.
            ASTNode expression = ParseExpression();

            // AST 노드를 생성하고 채웁니다.
            var assignmentNode = new AssignmentNode
            {
                Target = new VariableNode { Name = variableName },
                Value = expression
            };

            return assignmentNode;
        }

        private VariableNode ParseVariable()
        {
            var token = Match(Token.TokenType.Identifier);
            if (token == null)
                return null;

            Match(Token.TokenType.Keyword, "AS");
            var type = Match(Token.TokenType.Keyword);

            var variableName = token.Value;

            return new VariableNode { Name = variableName, Type = type?.Value };
        }

        private Token LookAhead()
        {
            if (_position >= _tokens.Count) 
                return null;

            return _tokens[_position];
        }

        private Token Match(Token.TokenType type, string value = null, int beforePosition = -1)
        {
            if (beforePosition > 0)
            {
                return _tokens[_position - beforePosition];
            }

            if (LookAhead().Type != type || (value != null && LookAhead().Value != value))
                return null;
            
            return _tokens[_position++];
        }

        public string Transpile(ASTNode ast, string funcName="")
        {
            StringBuilder sb = new StringBuilder();

            if (ast is SubroutineNode subProcedure)
            {
                string parameters = string.Join(", ", subProcedure.Parameters.ConvertAll(p => $"{p.Type} {p.Name}"));
                sb.AppendLine($"public void {subProcedure.Name}({parameters})");
                sb.AppendLine("{");
                foreach (var stmt in subProcedure.Body)
                {
                    sb.Append(Transpile(stmt));
                }
                sb.AppendLine();
                sb.AppendLine("}");
            }
            else if (ast is VariableNode dimStatement)
            {
                sb.AppendLine($"    {DataTypeMappings[dimStatement.Type]} {dimStatement.Name};");
            }
            else if (ast is AssignmentNode assignmentStatement)
            {

                if (funcName == assignmentStatement.Target.Name)
                {
                    sb.AppendLine();
                    sb.Append($"    return {ConvertExpressionToCSharp(assignmentStatement.Value)};");
                }
                else
                {
                    sb.Append($"    {assignmentStatement.Target.Name}");
                    sb.Append(" = ");
                    sb.Append(ConvertExpressionToCSharp(assignmentStatement.Value));
                    sb.Append(";");
                }
            }
            else if (ast is FunctionNode function)
            {
                string parameters = string.Join(", ", function.Parameters.ConvertAll(p => $"{DataTypeMappings[p.Type]} {p.Name}"));
                sb.AppendLine($"public {DataTypeMappings[function.Type]} {function.Name}({parameters})");
                sb.AppendLine("{");

                foreach (var stmt in function.Body)
                {
                    sb.Append(Transpile(stmt, function.Name));
                }

                sb.AppendLine();
                sb.AppendLine("}");
            }
            else if (ast is IfNode ifStatement)
            {
                 
            }

            return sb.ToString();
        }

        public string ConvertExpressionToCSharp(ASTNode node)
        {
            StringBuilder csharpCode = new StringBuilder();

            if (node is LiteralNode literalNode)
            {
                csharpCode.Append(literalNode.Value);
            }
            else if (node is VariableNode variableNode)
            {
                csharpCode.Append(variableNode.Name);
            }
            else if (node is BinaryOperationNode binaryOperationNode)
            {
                csharpCode.Append("(");
                csharpCode.Append(ConvertExpressionToCSharp(binaryOperationNode.Left));
                csharpCode.Append(" ");
                csharpCode.Append(binaryOperationNode.Operator);
                csharpCode.Append(" ");
                csharpCode.Append(ConvertExpressionToCSharp(binaryOperationNode.Right));
                csharpCode.Append(")");
            }
            // ... 기타 노드 타입에 대한 처리 ...

            return csharpCode.ToString();
        }

        private ASTNode ParseExpression()
        {
            // 첫 번째 피연산자 파싱
            ASTNode left = ParseOperand();

            while (LookAhead().Type == Token.TokenType.Operator)
            {
                // 연산자 파싱
                var op = Match(Token.TokenType.Operator).Value;

                // 두 번째 피연산자 파싱
                ASTNode right = ParseOperand();

                // 이항 연산자 노드 생성
                left = new BinaryOperationNode { Left = left, Right = right, Operator = op };
            }

            return left;
        }

        private ASTNode ParseOperand()
        {
            var token = LookAhead();
            switch (token.Type)
            {
                case Token.TokenType.Literal:
                    Match(Token.TokenType.Literal);
                    return new LiteralNode { Value = token.Value };
                case Token.TokenType.Identifier:
                    Match(Token.TokenType.Identifier);
                    return new VariableNode { Name = token.Value };

                default:
                    throw new Exception($"Invalid operand : {token.ToString()}");
            }
        }
    }
}