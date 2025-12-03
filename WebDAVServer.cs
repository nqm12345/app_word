using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using Newtonsoft.Json;

namespace WordWebDAV
{
    public class AppConfig
    {
        public int Port { get; set; } = 1900;
        public string CompanyApiUrl { get; set; } = "";
        public string ApiEndpoint { get; set; } = "/api/files/download";
    }

    public class WebDAVServer
    {
        private readonly HttpListener _listener;
        private readonly AppConfig _config;
        private readonly HttpClient _httpClient;
        private bool _isRunning;
        private Task? _serverTask;

        public event Action<string>? OnLog;
        public bool IsRunning => _isRunning;

        private string GetMimeType(string filename)
        {
            var ext = Path.GetExtension(filename).ToLowerInvariant();
            return ext switch
            {
                // Word
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".doc" => "application/msword",
                ".docm" => "application/vnd.ms-word.document.macroEnabled.12",
                ".rtf" => "application/rtf",
                // Excel
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".xls" => "application/vnd.ms-excel",
                ".xlsm" => "application/vnd.ms-excel.sheet.macroEnabled.12",
                ".csv" => "text/csv",
                // PowerPoint
                ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                ".ppt" => "application/vnd.ms-powerpoint",
                ".pptm" => "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
                // Visio
                ".vsdx" => "application/vnd.ms-visio.drawing",
                ".vsd" => "application/vnd.visio",
                // Project
                ".mpp" => "application/vnd.ms-project",
                // Default
                _ => "application/octet-stream"
            };
        }

        public WebDAVServer(AppConfig config)
        {
            _config = config;
            _listener = new HttpListener();
            _listener.Prefixes.Add($"http://localhost:{_config.Port}/");
            _listener.Prefixes.Add($"http://127.0.0.1:{_config.Port}/");
            
            var handler = new HttpClientHandler
            {
                ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true
            };
            _httpClient = new HttpClient(handler);
            _httpClient.Timeout = TimeSpan.FromMinutes(10);
        }

        public bool Start()
        {
            if (_isRunning) return true;

            try
            {
                _listener.Start();
                _isRunning = true;
                Log($"✅ Đã sẵn sàng trên cổng {_config.Port}");
                Log($"🌐 API: {_config.CompanyApiUrl}");

                _serverTask = Task.Run(async () =>
                {
                    while (_isRunning)
                    {
                        try
                        {
                            var context = await _listener.GetContextAsync();
                            _ = ProcessRequestAsync(context);
                        }
                        catch (Exception ex)
                        {
                            if (_isRunning)
                                Log($"   ⚠️ Lỗi xử lý: {ex.Message}");
                        }
                    }
                });
                return true;
            }
            catch (Exception ex)
            {
                _isRunning = false;
                
                // Translate error message to Vietnamese
                string errorMsg;
                if (ex.Message.Contains("conflicts with an existing"))
                {
                    errorMsg = $"❌ KHÔNG THỂ KHỞI ĐỘNG: Port {_config.Port} đang bị sử dụng bởi ứng dụng khác. Vui lòng đóng ứng dụng trùng cổng hoặc đổi port trong config.json";
                }
                else if (ex.Message.Contains("Access is denied"))
                {
                    errorMsg = "❌ KHÔNG THỂ KHỞI ĐỘNG: Không có quyền truy cập. Vui lòng chạy với quyền Administrator";
                }
                else
                {
                    errorMsg = $"❌ KHÔNG THỂ KHỞI ĐỘNG: {ex.Message}";
                }
                
                Log(errorMsg);
                return false;
            }
        }

        public void Stop()
        {
            _isRunning = false;
            try
            {
                _listener.Stop();
                _listener.Close();
                _httpClient.Dispose();
                Log("⏸️ Đã tạm dừng");
            }
            catch { }
        }

        private async Task ProcessRequestAsync(HttpListenerContext context)
        {
            var request = context.Request;
            var response = context.Response;

            try
            {
                var path = request.Url?.AbsolutePath ?? "";
                var method = request.HttpMethod;

                Log($"📥 [{method}] {path}");

                // Parse /files/:id/:userId/:object_type/:object_id/:edit_file_id/:filename
                if (path.StartsWith("/files/"))
                {
                    var parts = path.Substring(7).Split('/', 6);
                    if (parts.Length >= 1)
                    {
                        var id = parts[0];
                        var userId = parts.Length > 1 ? parts[1] : "";
                        var objectType = parts.Length > 2 ? parts[2] : "";
                        var objectId = parts.Length > 3 ? parts[3] : "";
                        var editFileId = parts.Length > 4 ? parts[4] : "";
                        var filename = parts.Length > 5 ? Uri.UnescapeDataString(parts[5]) : "document.docx";

                        switch (method)
                        {
                            case "GET":
                                await HandleGetAsync(id, userId, objectType, objectId, editFileId, filename, response);
                                break;
                            case "PUT":
                                await HandlePutAsync(id, userId, objectType, objectId, editFileId, filename, request, response);
                                break;
                            case "HEAD":
                                await HandleHeadAsync(id, userId, objectType, objectId, editFileId, filename, response);
                                break;
                            case "OPTIONS":
                                HandleOptions(response);
                                break;
                            case "PROPFIND":
                                HandlePropfind(id, filename, response);
                                break;
                            case "LOCK":
                                HandleLock(response);
                                break;
                            case "UNLOCK":
                                HandleUnlock(response);
                                break;
                            default:
                                response.StatusCode = 405;
                                break;
                        }
                        return;
                    }
                }

                // Status page
                if (path == "/" || path == "")
                {
                    var html = $@"
                        <html>
                        <head><title>Trình chỉnh sửa Word</title></head>
                        <body style='font-family:Arial;padding:20px;'>
                            <h1>✅ Trình chỉnh sửa Word đang hoạt động</h1>
                            <p>Cổng: {_config.Port}</p>
                            <p>API: {_config.CompanyApiUrl}</p>
                        </body>
                        </html>";
                    var buffer = Encoding.UTF8.GetBytes(html);
                    response.ContentType = "text/html; charset=utf-8";
                    await response.OutputStream.WriteAsync(buffer);
                }
                else
                {
                    response.StatusCode = 404;
                }
            }
            catch (Exception ex)
            {
                string connError = ex.Message;
                if (connError.Contains("No connection") || connError.Contains("Unable to connect"))
                    connError = "Không thể kết nối đến server. Kiểm tra kết nối Internet";
                else if (connError.Contains("timed out") || connError.Contains("Timeout"))
                    connError = "Hết thời gian chờ. Server phản hồi quá chậm";
                Log($"   ⚠️ LỖI KẾT NỐI: {connError}");
                response.StatusCode = 500;
            }
            finally
            {
                try { response.Close(); } catch { }
            }
        }

        private async Task HandleGetAsync(string id, string userId, string objectType, string objectId, string editFileId, string filename, HttpListenerResponse response)
        {
            var apiUrl = $"{_config.CompanyApiUrl}{_config.ApiEndpoint}/{id}/{userId}/{objectType}/{objectId}/{editFileId}/{Uri.EscapeDataString(filename)}";
            Log($"   → GET từ API...");

            var apiResponse = await _httpClient.GetAsync(apiUrl);
            
            if (!apiResponse.IsSuccessStatusCode)
            {
                string getError = apiResponse.StatusCode switch
                {
                    System.Net.HttpStatusCode.NotFound => "File không tồn tại trên server",
                    System.Net.HttpStatusCode.Unauthorized => "Không có quyền truy cập. Vui lòng đăng nhập lại",
                    System.Net.HttpStatusCode.Forbidden => "Bị từ chối truy cập file này",
                    System.Net.HttpStatusCode.InternalServerError => "Lỗi server. Vui lòng thử lại sau",
                    _ => "Không thể tải file từ server"
                };
                Log($"   ❌ LỖI MỞ FILE: {getError}");
                response.StatusCode = (int)apiResponse.StatusCode;
                return;
            }

            var content = await apiResponse.Content.ReadAsByteArrayAsync();
            Log($"   ✅ GET THÀNH CÔNG: {content.Length} bytes - File đã tải về Word");

            response.ContentType = GetMimeType(filename);
            response.AddHeader("Content-Disposition", $"inline; filename=\"{filename}\"");
            response.ContentLength64 = content.Length;
            await response.OutputStream.WriteAsync(content);
        }

        private async Task HandlePutAsync(string id, string userId, string objectType, string objectId, string editFileId, string filename, HttpListenerRequest request, HttpListenerResponse response)
        {
            // Read binary from Word
            using var ms = new MemoryStream();
            await request.InputStream.CopyToAsync(ms);
            var fileBytes = ms.ToArray();
            
            Log($"   📤 Nhận {fileBytes.Length} bytes từ Word");

            var apiUrl = $"{_config.CompanyApiUrl}{_config.ApiEndpoint}/{id}/{userId}/{objectType}/{objectId}/{editFileId}/{Uri.EscapeDataString(filename)}";
            Log($"   → PUT lên API...");

            // Send as multipart/form-data (Company API expects this)
            using var formContent = new MultipartFormDataContent();
            var fileContent = new ByteArrayContent(fileBytes);
            fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(GetMimeType(filename));
            formContent.Add(fileContent, "file", filename);

            var apiResponse = await _httpClient.PutAsync(apiUrl, formContent);
            
            if (!apiResponse.IsSuccessStatusCode)
            {
                var errorText = await apiResponse.Content.ReadAsStringAsync();
                string putError = apiResponse.StatusCode switch
                {
                    System.Net.HttpStatusCode.Unauthorized => "Không có quyền lưu. Vui lòng đăng nhập lại",
                    System.Net.HttpStatusCode.Forbidden => "Bị từ chối lưu file này",
                    System.Net.HttpStatusCode.RequestEntityTooLarge => "File quá lớn, không thể lưu",
                    System.Net.HttpStatusCode.InternalServerError => "Lỗi server. Vui lòng thử lại sau",
                    _ => "Không thể lưu file lên server. Kiểm tra kết nối Internet"
                };
                Log($"   ❌ LỖI LƯU FILE: {putError}");
                response.StatusCode = (int)apiResponse.StatusCode;
                return;
            }

            var result = await apiResponse.Content.ReadAsStringAsync();
            Log($"   ✅ LƯU THÀNH CÔNG: File đã được lưu lên server");

            response.ContentType = "application/json";
            var buffer = Encoding.UTF8.GetBytes(result);
            await response.OutputStream.WriteAsync(buffer);
        }

        private async Task HandleHeadAsync(string id, string userId, string objectType, string objectId, string editFileId, string filename, HttpListenerResponse response)
        {
            var apiUrl = $"{_config.CompanyApiUrl}{_config.ApiEndpoint}/{id}/{userId}/{objectType}/{objectId}/{editFileId}/{Uri.EscapeDataString(filename)}";
            
            var request = new HttpRequestMessage(HttpMethod.Head, apiUrl);
            var apiResponse = await _httpClient.SendAsync(request);
            
            response.StatusCode = (int)apiResponse.StatusCode;
            response.ContentType = GetMimeType(filename);
        }

        private void HandleOptions(HttpListenerResponse response)
        {
            response.AddHeader("Allow", "GET, PUT, HEAD, OPTIONS, PROPFIND, LOCK, UNLOCK");
            response.AddHeader("DAV", "1, 2");
            response.AddHeader("MS-Author-Via", "DAV");
            response.StatusCode = 200;
        }

        private void HandlePropfind(string id, string filename, HttpListenerResponse response)
        {
            var xml = $@"<?xml version=""1.0"" encoding=""utf-8""?>
                <D:multistatus xmlns:D=""DAV:"">
                    <D:response>
                        <D:href>/files/{id}/{filename}</D:href>
                        <D:propstat>
                            <D:prop>
                                <D:displayname>{filename}</D:displayname>
                                <D:getcontenttype>application/vnd.openxmlformats-officedocument.wordprocessingml.document</D:getcontenttype>
                                <D:resourcetype/>
                            </D:prop>
                            <D:status>HTTP/1.1 200 OK</D:status>
                        </D:propstat>
                    </D:response>
                </D:multistatus>";
            
            response.ContentType = "application/xml; charset=utf-8";
            response.StatusCode = 207;
            var buffer = Encoding.UTF8.GetBytes(xml);
            response.OutputStream.Write(buffer, 0, buffer.Length);
        }

        private void HandleLock(HttpListenerResponse response)
        {
            var xml = $@"<?xml version=""1.0"" encoding=""utf-8""?>
                <D:prop xmlns:D=""DAV:"">
                    <D:lockdiscovery>
                        <D:activelock>
                            <D:locktoken>
                                <D:href>opaquelocktoken:{Guid.NewGuid()}</D:href>
                            </D:locktoken>
                            <D:timeout>Second-3600</D:timeout>
                        </D:activelock>
                    </D:lockdiscovery>
                </D:prop>";
            
            response.ContentType = "application/xml; charset=utf-8";
            response.StatusCode = 200;
            var buffer = Encoding.UTF8.GetBytes(xml);
            response.OutputStream.Write(buffer, 0, buffer.Length);
        }

        private void HandleUnlock(HttpListenerResponse response)
        {
            response.StatusCode = 204;
        }

        private void Log(string message)
        {
            var time = DateTime.Now.ToString("HH:mm:ss");
            OnLog?.Invoke($"[{time}] {message}");
        }
    }
}
