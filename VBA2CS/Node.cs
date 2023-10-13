using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBA2CS
{
    abstract class AstNode { }

    class AstVariableDeclaration : AstNode
    {
        public string VariableName { get; set; }
        public string VariableType { get; set; }
    }

    class AstAssignment : AstNode
    {
        public string VariableName { get; set; }
        public AstExpression Expression { get; set; }
    }

    abstract class AstExpression : AstNode { }

    class AstLiteral : AstExpression
    {
        public string Value { get; set; }
    }

    class AstBinaryOperation : AstExpression
    {
        public AstExpression Left { get; set; }
        public string Operator { get; set; }
        public AstExpression Right { get; set; }
    }

    class AstFunctionDeclaration : AstNode
    {
        public string FunctionName { get; set; }
        public string ReturnType { get; set; }
        public List<AstParameter> Parameters { get; set; }
        public List<AstNode> Body { get; set; }
    }

    class AstIfStatement : AstNode
    {
        public AstExpression Condition { get; set; }
        public List<AstNode> ThenBody { get; set; }
        public List<AstNode> ElseBody { get; set; }
    }

    class AstForLoop : AstNode
    {
        public AstAssignment Initialization { get; set; }
        public AstExpression Condition { get; set; }
        public AstAssignment Update { get; set; }
        public List<AstNode> Body { get; set; }
    }

    class AstArrayDeclaration : AstNode
    {
        public string ArrayName { get; set; }
        public string ElementType { get; set; }
        public int Dimensions { get; set; }
        public List<AstExpression> InitialValues { get; set; }
    }

    // 사용자 정의 타입 노드
    class AstUserDefinedType : AstNode
    {
        public string TypeName { get; set; }
        public List<AstVariableDeclaration> Fields { get; set; }
    }

    // 열거형 노드
    class AstEnumDeclaration : AstNode
    {
        public string EnumName { get; set; }
        public List<AstEnumMember> Members { get; set; }
    }

    class AstEnumMember : AstNode
    {
        public string MemberName { get; set; }
        public AstExpression Value { get; set; }
    }

    class AstProcedureCall : AstNode
    {
        public string ProcedureName { get; set; }
        public List<AstExpression> Arguments { get; set; }
    }

    // Select Case 문 노드
    class AstSelectCase : AstNode
    {
        public AstExpression Condition { get; set; }
        public List<AstCaseBlock> Cases { get; set; }
    }

    class AstCaseBlock : AstNode
    {
        public AstExpression CaseValue { get; set; }
        public List<AstNode> Body { get; set; }
    }

    // 오류 처리 노드
    class AstOnError : AstNode
    {
        public string ErrorHandler { get; set; }
    }

    // With 문 노드
    class AstWithStatement : AstNode
    {
        public AstExpression Object { get; set; }
        public List<AstNode> Body { get; set; }
    }

    // Do Loop 문 노드
    class AstDoLoop : AstNode
    {
        public AstExpression Condition { get; set; }
        public List<AstNode> Body { get; set; }
    }

    // ByVal, ByRef 매개변수 노드
    class AstParameter : AstNode
    {
        public string ParameterName { get; set; }
        public string ParameterType { get; set; }
        public bool IsByVal { get; set; }
        public bool IsOptional { get; set; }
        public AstExpression DefaultValue { get; set; }
    }

    // Variant 타입 노드
    class AstVariant : AstNode
    {
        public AstExpression Value { get; set; }
    }

    // Exit 문 노드 (Exit For, Exit Do, Exit Sub, Exit Function 등)
    class AstExitStatement : AstNode
    {
        public string ExitType { get; set; }
    }

    // GoTo 문 노드
    class AstGoToStatement : AstNode
    {
        public string Label { get; set; }
    }

    // 레이블 노드
    class AstLabel : AstNode
    {
        public string LabelName { get; set; }
    }
}
