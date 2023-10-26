using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace VBA2CS
{
    public abstract class ASTNode
    {
        public List<ASTNode> Children { get; } = new List<ASTNode>();
    }

    public class FunctionNode : ASTNode
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public List<VariableNode> Parameters { get; set; } = new List<VariableNode>();
        public List<ASTNode> Body { get; set; } = new List<ASTNode>();
    }

    public class SubroutineNode : ASTNode
    {
        public string Name { get; set; }
        public List<VariableNode> Parameters { get; } = new List<VariableNode>();
        public List<ASTNode> Body { get; set; } = new List<ASTNode>();
    }

    public class VariableNode : ASTNode
    {
        public string Name { get; set; }
        public string Type { get; set; }
    }

    public class ConstantNode : ASTNode
    {
        public string Name { get; set; }
        public Token.TokenType Type { get; set; }
        public ASTNode Value { get; set; }
    }

    // 연산자와 표현식
    public class BinaryOperationNode : ASTNode
    {
        public string Operator { get; set; }
        public ASTNode Left { get; set; }
        public ASTNode Right { get; set; }
    }

    public class UnaryOperationNode : ASTNode
    {
        public Token.TokenType Operator { get; set; }
        public ASTNode Operand { get; set; }
    }

    public class LiteralNode : ASTNode
    {
        public string Value { get; set; }
    }

    // 제어 구조
    public class IfNode : ASTNode
    {
        public ASTNode Condition { get; set; }
        public List<ASTNode> TrueBranch { get; set; } = new List<ASTNode>();
        public List<ASTNode> FalseBranch { get; set; } = new List<ASTNode>();
    }

    public class ForLoopNode : ASTNode
    {
        public VariableNode Iterator { get; set; }
        public ASTNode Start { get; set; }
        public ASTNode End { get; set; }
        public ASTNode Step { get; set; } // 생략 가능
    }

    public class WhileLoopNode : ASTNode
    {
        public ASTNode Condition { get; set; }
    }

    public class DoLoopNode : ASTNode
    {
        public ASTNode Condition { get; set; }
        public bool Until { get; set; } // Do ... Loop Until 형태인 경우 true
    }

    public class CommentNode : ASTNode
    {
        public string Value { get; set; }
    }

    // 기본 문장 노드
    public class StatementNode : ASTNode { }

    // 변수 할당
    public class AssignmentNode : StatementNode
    {
        public VariableNode Target { get; set; }
        public ASTNode Value { get; set; }
    }

    // 함수 호출
    public class FunctionCallNode : StatementNode
    {
        public string FunctionName { get; set; }
        public List<ASTNode> Arguments { get; } = new List<ASTNode>();
    }

    // 배열
    public class ArrayNode : ASTNode
    {
        public string Name { get; set; }
        public List<ASTNode> Indices { get; } = new List<ASTNode>();
    }

    // Select Case 문
    public class SelectCaseNode : StatementNode
    {
        public ASTNode TestExpression { get; set; }
        public List<CaseNode> Cases { get; } = new List<CaseNode>();
    }

    public class CaseNode : ASTNode
    {
        public List<ASTNode> Conditions { get; } = new List<ASTNode>();
        public List<StatementNode> Statements { get; } = new List<StatementNode>();
    }

    // With 문
    public class WithNode : StatementNode
    {
        public ASTNode Target { get; set; }
        public List<StatementNode> Statements { get; } = new List<StatementNode>();
    }

    // On Error 문
    public class OnErrorNode : StatementNode
    {
        public enum ErrorMode
        {
            Next,
            GoTo,
            ResumeNext
        }

        public ErrorMode Mode { get; set; }
        public string Label { get; set; } // GoTo 레이블
    }

    // GoTo 문
    public class GoToNode : StatementNode
    {
        public string Label { get; set; }
    }

    // 레이블
    public class LabelNode : StatementNode
    {
        public string Name { get; set; }
    }
}
