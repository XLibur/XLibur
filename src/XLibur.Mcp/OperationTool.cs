using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using XLibur.Operations;

namespace XLibur.Mcp;

internal sealed class OperationTool : McpServerTool
{
    private const string FormatArgument = "format";

    private readonly IOperation _operation;
    private readonly OperationContext _context;
    private readonly SemaphoreSlim _gate;

    public OperationTool(IOperation operation, OperationContext context, SemaphoreSlim gate)
    {
        _operation = operation;
        _context = context;
        _gate = gate;
        var readOnly = operation.Effect == OperationEffect.ReadOnly;
        ProtocolTool = new Tool
        {
            Name = operation.Name,
            Description = operation.Description,
            InputSchema = Schema(operation.RequestType),
            Annotations = new ToolAnnotations
            {
                ReadOnlyHint = readOnly,
                DestructiveHint = !readOnly,
                IdempotentHint = readOnly,
                OpenWorldHint = false,
            },
        };
    }

    public override Tool ProtocolTool { get; }

    public override IReadOnlyList<object> Metadata => [];

    public override ValueTask<CallToolResult> InvokeAsync(RequestContext<CallToolRequestParams> request,
        CancellationToken cancellationToken = default)
        => CallAsync(request.Params?.Arguments, cancellationToken);

    /// <summary>
    /// Binds the call's arguments to the request type, then runs the operation under the shared gate. A
    /// <c>format</c> argument picks text or JSON; an argument the request type does not have is an error.
    /// </summary>
    internal async ValueTask<CallToolResult> CallAsync(IEnumerable<KeyValuePair<string, JsonElement>>? callArguments,
        CancellationToken cancellationToken)
    {
        var arguments = new JsonObject();
        var format = ResultFormat.Text;
        foreach (var (name, value) in callArguments ?? [])
        {
            if (name == FormatArgument)
            {
                if (!Enum.TryParse(value.ValueKind == JsonValueKind.String ? value.GetString() : null, true, out format))
                    return Error(new OperationError(ErrorKind.BadInput, "bad_arguments", "format is \"text\" or \"json\"."), ResultFormat.Text);
                continue;
            }

            arguments[name] = JsonNode.Parse(value.GetRawText());
        }

        object parsed;
        try
        {
            parsed = arguments.Deserialize(_operation.RequestType, Rendering.JsonOptions)
                     ?? throw new JsonException("The arguments are empty.");
        }
        catch (JsonException e)
        {
            return Error(new OperationError(ErrorKind.BadInput, "bad_arguments", e.Message), format);
        }

        await _gate.WaitAsync(cancellationToken);
        try
        {
            var (outcome, text) = await Task.Run(() =>
            {
                var result = OperationRunner.Run(_operation, parsed, _context, format, out var rendered);
                return (result, rendered);
            }, cancellationToken);
            return Result(text, outcome.Error is not null);
        }
        finally
        {
            _gate.Release();
        }
    }

    private static CallToolResult Error(OperationError error, ResultFormat format)
        => Result(OperationOutcome.Failure(error).Render(format), isError: true);

    private static CallToolResult Result(string text, bool isError) => new()
    {
        Content = [new TextContentBlock { Text = text }],
        IsError = isError,
    };

    private static JsonElement Schema(Type requestType)
    {
        var schema = JsonNode.Parse(AIJsonUtilities.CreateJsonSchema(requestType, serializerOptions: Rendering.JsonOptions).GetRawText())!.AsObject();
        var properties = schema["properties"]?.AsObject() ?? new JsonObject();
        properties[FormatArgument] = new JsonObject
        {
            ["type"] = "string",
            ["enum"] = new JsonArray("text", "json"),
            ["description"] = "text (default) is compact TSV-style output; json is the full structured result.",
        };
        schema["properties"] = properties;
        return JsonSerializer.SerializeToElement(schema);
    }
}
