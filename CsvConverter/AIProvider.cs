using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

namespace CsvConverter
{
    public interface IAIProvider
    {
        Task<string> GenerateAsync(string prompt, CancellationToken token);
    }

    public class OllamaProvider : IAIProvider
    {
        private readonly HttpClient _httpClient;
        private readonly string _baseUrl;
        private readonly string _model;

        public OllamaProvider(string baseUrl, string model = "llama2", int timeoutSeconds = 120)
        {
            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(timeoutSeconds) };
            _baseUrl = baseUrl.TrimEnd('/');
            _model = model;
        }

        public async Task<string> GenerateAsync(string prompt, CancellationToken token)
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            AIProviderFactory.LogToFile($"Generating AI response for prompt: '{prompt}' (model: {_model})");
            try
            {
                var requestBody = new
                {
                    model = _model,
                    prompt = prompt,
                    stream = false,
                    options = new
                    {
                        num_predict = 100,
                        temperature = 0.3
                    }
                };

                var json = JsonSerializer.Serialize(requestBody);
                AIProviderFactory.LogToFile($"Request JSON: {json}");
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync($"{_baseUrl}/api/generate", content, token);
                response.EnsureSuccessStatusCode();

                var responseJson = await response.Content.ReadAsStringAsync(token);
                // Log full response (truncated if too long)
                var responseLog = responseJson.Length > 2000 ? responseJson.Substring(0, 2000) + "..." : responseJson;
                AIProviderFactory.LogToFile($"Response JSON: {responseLog}");
                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };
                var result = JsonSerializer.Deserialize<OllamaResponse>(responseJson, options);
                var aiResponse = result?.Response ?? string.Empty;
                stopwatch.Stop();
                AIProviderFactory.LogToFile($"AI response received for prompt: '{prompt}' -> '{aiResponse}' (took {stopwatch.ElapsedMilliseconds} ms)");
                return aiResponse;
            }
            catch (HttpRequestException ex)
            {
                stopwatch.Stop();
                AIProviderFactory.LogToFile($"HTTP error for prompt '{prompt}' after {stopwatch.ElapsedMilliseconds} ms: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Error connecting to Ollama at {_baseUrl}: {ex.Message}");
                throw;
            }
            catch (TaskCanceledException ex)
            {
                stopwatch.Stop();
                AIProviderFactory.LogToFile($"Timeout for prompt '{prompt}' after {stopwatch.ElapsedMilliseconds} ms: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Ollama request timeout: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                AIProviderFactory.LogToFile($"Unexpected error for prompt '{prompt}' after {stopwatch.ElapsedMilliseconds} ms: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Ollama error: {ex.Message}");
                throw;
            }
        }

        private class OllamaResponse
        {
            [JsonPropertyName("response")]
            public string Response { get; set; } = string.Empty;
        }
    }

    public static class AIProviderFactory
    {
        /// <summary>
        /// Директория для лог-файла AI провайдера. Если null, используется %APPDATA%\CsvConverter\logs.
        /// </summary>
        public static string? LogDirectory { get; set; }

        public static void LogToFile(string message)
        {
            try
            {
                string logDir;
                if (LogDirectory != null && Directory.Exists(LogDirectory))
                {
                    logDir = LogDirectory;
                }
                else
                {
                    logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CsvConverter", "logs");
                }
                Directory.CreateDirectory(logDir);
                var logFile = Path.Combine(logDir, "ai-provider.log");
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                File.AppendAllText(logFile, $"[{timestamp}] {message}{Environment.NewLine}");
            }
            catch
            {
                // Ignore logging errors
            }
        }

        public static async Task<IAIProvider?> CreateProviderAsync(AIConfig config, CancellationToken token = default)
        {
            try
            {
                // Log to file for debugging
                LogToFile($"AIProviderFactory.CreateProviderAsync called: Name='{config.Name}', URL='{config.Url}', Model='{config.Model}'");
                
                if (config.Name?.ToLower() == "ollama" && !string.IsNullOrEmpty(config.Url))
                {
                    var model = config.Model;
                    
                    // Log the raw config values for debugging
                    LogToFile($"AI config raw: TimeoutSeconds={config.TimeoutSeconds}, Url={config.Url}, Model={model}");
                    
                    // If model is not specified, try to get the first available model from Ollama
                    if (string.IsNullOrEmpty(model))
                    {
                        LogToFile("Model not specified, attempting to auto-detect...");
                        var detectTimeout = config.TimeoutSeconds > 0 ? config.TimeoutSeconds : 30;
                        model = await GetFirstAvailableModelAsync(config.Url, detectTimeout, token);
                        LogToFile($"Auto-detected model: {model ?? "(null)"}");
                    }

                    if (string.IsNullOrEmpty(model))
                    {
                        model = "llama2"; // Final fallback
                        LogToFile($"Using fallback model: {model}");
                    }

                    var timeoutSeconds = config.TimeoutSeconds > 0 ? config.TimeoutSeconds : 120;
                    System.Diagnostics.Debug.WriteLine($"Creating OllamaProvider: URL={config.Url}, Model={model}, Timeout={timeoutSeconds}s");
                    LogToFile($"Creating OllamaProvider: URL={config.Url}, Model={model}, Timeout={timeoutSeconds}s");
                    return new OllamaProvider(config.Url, model, timeoutSeconds);
                }

                System.Diagnostics.Debug.WriteLine($"AIProviderFactory: Name='{config.Name}', URL='{config.Url}' - provider not recognized");
                LogToFile($"AIProviderFactory: Name='{config.Name}', URL='{config.Url}' - provider not recognized");
                return null;
            }
            catch (Exception ex)
            {
                LogToFile($"AIProviderFactory exception: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"AIProviderFactory exception: {ex.Message}");
                return null;
            }
        }

        private static async Task<string?> GetFirstAvailableModelAsync(string baseUrl, int timeoutSeconds = 30, CancellationToken token = default)
        {
            try
            {
                using var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(timeoutSeconds) };
                var response = await httpClient.GetAsync($"{baseUrl.TrimEnd('/')}/api/tags", token);
                
                if (!response.IsSuccessStatusCode)
                    return null;

                var responseJson = await response.Content.ReadAsStringAsync(token);
                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };
                var tagsResponse = JsonSerializer.Deserialize<OllamaTagsResponse>(responseJson, options);
                if (tagsResponse?.Models != null && tagsResponse.Models.Length > 0)
                {
                    var modelName = tagsResponse.Models[0].Name;
                    System.Diagnostics.Debug.WriteLine($"Auto-detected model from Ollama: {modelName}");
                    return modelName;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to auto-detect Ollama models: {ex.Message}");
            }

            return null;
        }

        private class OllamaTagsResponse
        {
            [JsonPropertyName("models")]
            public OllamaModel[] Models { get; set; } = Array.Empty<OllamaModel>();
        }

        private class OllamaModel
        {
            [JsonPropertyName("name")]
            public string Name { get; set; } = string.Empty;
        }
    }
}