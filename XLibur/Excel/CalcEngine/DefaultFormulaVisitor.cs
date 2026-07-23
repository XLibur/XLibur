using System.Collections.Generic;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// A default visitor that copies a formula.
/// </summary>
internal class DefaultFormulaVisitor<TContext> : IFormulaVisitor<TContext, AstNode>
{
    public virtual AstNode Visit(TContext context, UnaryNode node)
    {
        var acceptedArgument = (ValueNode)node.Expression.Accept(context, this);
        return !ReferenceEquals(acceptedArgument, node.Expression)
            ? new UnaryNode(node.Operation, acceptedArgument)
            : node;
    }

    public virtual AstNode Visit(TContext context, BinaryNode node)
    {
        var acceptedLeftArgument = (ValueNode)node.LeftExpression.Accept(context, this);
        var acceptedRightArgument = (ValueNode)node.RightExpression.Accept(context, this);
        return !ReferenceEquals(acceptedLeftArgument, node.LeftExpression) || !ReferenceEquals(acceptedRightArgument, node.RightExpression)
            ? new BinaryNode(node.Operation, acceptedLeftArgument, acceptedRightArgument)
            : node;
    }

    public virtual AstNode Visit(TContext context, FunctionNode node)
    {
        var parameters = node.Parameters;

        // Only materialize a new parameter list once a child actually changes; if every parameter
        // is returned unchanged, reuse the original node (matching the Unary/Binary visitors).
        List<ValueNode>? acceptedParameters = null;
        for (var i = 0; i < parameters.Count; i++)
        {
            var parameter = parameters[i];
            var acceptedParameter = (ValueNode)parameter.Accept(context, this);

            if (acceptedParameters is null && !ReferenceEquals(acceptedParameter, parameter))
            {
                acceptedParameters = new List<ValueNode>(parameters.Count);
                for (var j = 0; j < i; j++)
                    acceptedParameters.Add(parameters[j]);
            }

            acceptedParameters?.Add(acceptedParameter);
        }

        return acceptedParameters is not null
            ? new FunctionNode(node.Prefix, node.Name, acceptedParameters)
            : node;
    }

    public virtual AstNode Visit(TContext context, ScalarNode node) => node;

    public virtual AstNode Visit(TContext context, ArrayNode node) => node;

    public virtual AstNode Visit(TContext context, NotSupportedNode node) => node;

    public virtual AstNode Visit(TContext context, ReferenceNode node)
    {
        var acceptedPrefix = node.Prefix?.Accept(context, this);
        return !ReferenceEquals(acceptedPrefix, node.Prefix)
            ? new ReferenceNode((PrefixNode?)acceptedPrefix, node.ReferenceArea, node.IsA1)
            : node;
    }

    public virtual AstNode Visit(TContext context, NameNode node)
    {
        var acceptedPrefix = node.Prefix?.Accept(context, this);
        return !ReferenceEquals(acceptedPrefix, node.Prefix)
            ? new NameNode((PrefixNode?)acceptedPrefix, node.Name)
            : node;
    }

    public virtual AstNode Visit(TContext context, StructuredReferenceNode node) => node;

    public virtual AstNode Visit(TContext context, PrefixNode node)
    {
        var acceptedFile = node.File?.Accept(context, this);
        return !ReferenceEquals(acceptedFile, node.File)
            ? new PrefixNode((FileNode?)acceptedFile, node.Sheet, node.FirstSheet, node.LastSheet)
            : node;
    }

    public virtual AstNode Visit(TContext context, FileNode node) => node;
}
