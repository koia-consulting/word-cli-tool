using Amazon.Lambda.APIGatewayEvents;
using Amazon.Lambda.Core;
using HttpMultipartParser;
using System.Text;
using System.Text.Json;

[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace WordDocProcessor;

public class Function
{
    public async Task<APIGatewayProxyResponse> FunctionHandler(APIGatewayProxyRequest request, ILambdaContext context)
    {
        try
        {
            // Safely get Content-Type header
            var contentTypeHeader = request.Headers.TryGetValue("Content-Type", out var ct)
                ? ct
                : request.Headers.TryGetValue("content-type", out var ctLower)
                    ? ctLower
                    : throw new Exception("Missing Content-Type header");

            // Extract boundary manually
            var boundary = GetBoundary(contentTypeHeader);
            if (string.IsNullOrEmpty(boundary))
                throw new Exception("Missing boundary in Content-Type header.");

            // Decode the request body
            var bodyBytes = request.IsBase64Encoded
                ? Convert.FromBase64String(request.Body)
                : Encoding.UTF8.GetBytes(request.Body);

            using var stream = new MemoryStream(bodyBytes);

            // Parse multipart with boundary
            var parser = await MultipartFormDataParser.ParseAsync(stream, boundary);

            var file = parser.Files[0];
            using var reader = new StreamReader(file.Data);
            var fileContent = await reader.ReadToEndAsync();

            var changes = parser.GetParameterValue("changes");

            var result = new
            {
                fileName = file.FileName,
                fileContent = fileContent,
                metadata = changes
            };

            return new APIGatewayProxyResponse
            {
                StatusCode = 200,
                Body = JsonSerializer.Serialize(result),
                Headers = new Dictionary<string, string> { { "Content-Type", "application/json" } }
            };
        }
        catch (Exception ex)
        {
            return new APIGatewayProxyResponse
            {
                StatusCode = 500,
                Body = $"Error: {ex.Message}\n{ex.StackTrace}",
                Headers = new Dictionary<string, string> { { "Content-Type", "text/plain" } }
            };
        }
    }

    private string? GetBoundary(string contentType)
    {
        var boundaryKey = "boundary=";
        var boundaryIndex = contentType.IndexOf(boundaryKey, StringComparison.OrdinalIgnoreCase);
        return boundaryIndex >= 0 ? contentType.Substring(boundaryIndex + boundaryKey.Length).Trim('"') : null;
    }
}

