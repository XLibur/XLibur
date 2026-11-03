using System.Text;

namespace XLibur.Operations;

/// <summary>Whether an operation can change a file. Hosts gate and annotate on it.</summary>
public enum OperationEffect
{
    ReadOnly,
    Writes,
}

/// <summary>
/// One entry of the operation table. A host knows nothing of any single operation: it turns its own
/// input into an instance of <see cref="RequestType"/>, calls <see cref="Execute"/>, and renders
/// the result. The request type is the input schema both hosts derive from.
/// </summary>
public interface IOperation
{
    /// <summary>The tool name, in snake_case, such as <c>read_range</c>. The CLI command is the kebab-case form.</summary>
    string Name { get; }

    string Description { get; }

    OperationEffect Effect { get; }

    Type RequestType { get; }

    IOperationResult Execute(object request, OperationContext context);
}

public abstract class Operation<TRequest, TResult> : IOperation
    where TRequest : class
    where TResult : IOperationResult
{
    public abstract string Name { get; }

    public abstract string Description { get; }

    public abstract OperationEffect Effect { get; }

    public Type RequestType => typeof(TRequest);

    public abstract TResult Execute(TRequest request, OperationContext context);

    IOperationResult IOperation.Execute(object request, OperationContext context) => Execute((TRequest)request, context);
}

/// <summary>A result: plain data for JSON, plus a compact text form for an agent's context or a terminal.</summary>
public interface IOperationResult
{
    void WriteText(StringBuilder text);

    /// <summary>The process exit code of a CLI run that ends in this result. Non-zero says "differs" or "has issues".</summary>
    int ExitCode => 0;
}

/// <summary>Binds a request property to a positional CLI argument.</summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class ArgumentAttribute(int position) : Attribute
{
    public int Position { get; } = position;
}

public enum ResultFormat
{
    Text,
    Json,
}

/// <summary>The end of one run: a result or an error, never both.</summary>
public sealed class OperationOutcome
{
    private OperationOutcome(IOperationResult? result, OperationError? error)
    {
        Result = result;
        Error = error;
    }

    public IOperationResult? Result { get; }

    public OperationError? Error { get; }

    public int ExitCode => Error is null ? Result!.ExitCode : 2;

    public static OperationOutcome Success(IOperationResult result) => new(result, null);

    public static OperationOutcome Failure(OperationError error) => new(null, error);

    public string Render(ResultFormat format) => Error is not null
        ? Rendering.RenderError(Error, format)
        : Rendering.RenderResult(Result!, format);
}

public static class OperationRunner
{
    /// <summary>
    /// Runs one operation. An <see cref="OperationException"/> becomes its error; any other exception
    /// is a defect, reported with its type so it can be filed. A rendered result longer than
    /// <see cref="Limits.MaxResultChars"/> is replaced by a <c>result_too_large</c> error.
    /// </summary>
    public static OperationOutcome Run(IOperation operation, object request, OperationContext context, ResultFormat format,
        out string rendered)
    {
        OperationOutcome outcome;
        try
        {
            outcome = OperationOutcome.Success(operation.Execute(request, context));
            rendered = outcome.Render(format);
        }
        catch (OperationException e)
        {
            outcome = OperationOutcome.Failure(e.Error);
            rendered = outcome.Render(format);
        }
        catch (Exception e)
        {
            outcome = OperationOutcome.Failure(new OperationError(ErrorKind.Defect, "defect",
                $"{e.GetType().FullName}: {e.Message}"));
            rendered = outcome.Render(format);
        }

        if (rendered.Length <= context.Limits.MaxResultChars)
            return outcome;

        outcome = OperationOutcome.Failure(new OperationError(ErrorKind.BadInput, "result_too_large",
            $"The result is {rendered.Length:N0} characters; the limit is {context.Limits.MaxResultChars:N0}. Narrow the request."));
        rendered = outcome.Render(format);
        return outcome;
    }
}
