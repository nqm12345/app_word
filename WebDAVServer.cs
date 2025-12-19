using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace WordWebDAV
{
    public class AppConfig
    {
        public int Port { get; set; } = 1900;
        public string CompanyApiUrl { get; set; } = "";
        public string[] CompanyApiUrls { get; set; } = new string[0];
        public string ApiEndpoint { get; set; } = "/api/files/download";
        
        // Lấy danh sách URLs (hỗ trợ cả config cũ và mới)
        public string[] GetApiUrls()
        {
            if (CompanyApiUrls != null && CompanyApiUrls.Length > 0)
                return CompanyApiUrls;
            if (!string.IsNullOrEmpty(CompanyApiUrl))
                return new[] { CompanyApiUrl };
            return new[] { "https://administrator.lifetex.vn:316" };
        }
    }

    public class WebDAVServer
    {
        private readonly HttpListener _listener;
        private readonly AppConfig _config;
        private readonly HttpClient _httpClient;
        private bool _isRunning;
        private Task _serverTask;
        private string _fastestApiUrl; // Cache URL nhanh nhất
        // Bộ đệm log phục vụ qua HTTP (xem tại /logs)
        private readonly System.Collections.Concurrent.ConcurrentQueue<string> _logBuffer = new System.Collections.Concurrent.ConcurrentQueue<string>();
        private const int MaxLogLines = 500;
        // Lưu filename gốc để đảm bảo PUT dùng đúng tên file (key = path không có filename)
        private readonly System.Collections.Concurrent.ConcurrentDictionary<string, string> _originalFilenames = new System.Collections.Concurrent.ConcurrentDictionary<string, string>();

        public event Action<string> OnLog;
        public bool IsRunning => _isRunning;

        private void FindFastestServer()
        {
            var apiUrls = _config.GetApiUrls();
            Log($"🔍 Đang tìm server nhanh nhất trong {apiUrls.Length} server...");
            
            // Thử kết nối đến từng server để tìm server nhanh nhất
            string fastestUrl = "";
            long fastestTime = long.MaxValue;
            
            foreach (var url in apiUrls)
            {
                Log($"   → Kiểm tra: {url}");
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                
                try
                {
                    // Dùng HEAD request đến root - nhanh và nhẹ
                    using var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(3));
                    var request = new HttpRequestMessage(HttpMethod.Head, url);
                    var response = _httpClient.SendAsync(request, cts.Token).Result;
                    
                    stopwatch.Stop();
                    var elapsed = stopwatch.ElapsedMilliseconds;
                    
                    // Chỉ cần server phản hồi (bất kể status code)
                    Log($"      ✅ {url} phản hồi trong {elapsed}ms");
                    
                    if (elapsed < fastestTime)
                    {
                        fastestTime = elapsed;
                        fastestUrl = url;
                    }
                }
                catch (Exception ex)
                {
                    stopwatch.Stop();
                    string errMsg = ex.InnerException?.Message ?? ex.Message;
                    if (errMsg.Contains("No connection") || errMsg.Contains("Unable to connect"))
                        Log($"      ❌ {url} không kết nối được");
                    else if (errMsg.Contains("timed out") || errMsg.Contains("canceled"))
                        Log($"      ❌ {url} timeout (>3s)");
                    else
                        Log($"      ❌ {url} lỗi: {errMsg}");
                }
            }
            
            if (!string.IsNullOrEmpty(fastestUrl))
            {
                _fastestApiUrl = fastestUrl;
                Log($"   ✅ Server nhanh nhất: {_fastestApiUrl} ({fastestTime}ms)");
            }
            else
            {
                // Fallback - dùng server đầu tiên
                _fastestApiUrl = apiUrls.Length > 0 ? apiUrls[0] : "";
                Log($"   ⚠️ Không server nào phản hồi, dùng mặc định: {_fastestApiUrl}");
            }
        }

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
            
            // === TỐI ƯU TỐC ĐỘ ===
            
            // Tăng số kết nối đồng thời
            ServicePointManager.DefaultConnectionLimit = 100;
            ServicePointManager.Expect100Continue = false;
            ServicePointManager.UseNagleAlgorithm = false;
            
            // Bỏ qua SSL certificate validation cho Windows 7+
            ServicePointManager.ServerCertificateValidationCallback = 
                delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            
            // HttpClient với tối ưu cho .NET 4.5
            var handler = new HttpClientHandler
            {
                UseProxy = false, // Không dùng proxy = nhanh hơn
                AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate
            };
            _httpClient = new HttpClient(handler);
            _httpClient.Timeout = TimeSpan.FromMinutes(10);
            _httpClient.DefaultRequestHeaders.ConnectionClose = false; // Keep-alive
        }

        public bool Start()
        {
            if (_isRunning) return true;

            try
            {
                _listener.Start();
                _isRunning = true;
                Log($"✅ Đã sẵn sàng trên cổng {_config.Port}");
                
                // Tìm server nhanh nhất khi khởi động
                FindFastestServer();

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

                // Warm up - gửi nhiều request để "khởi động" server hoàn toàn
                Task.Run(async () =>
                {
                    try
                    {
                        await Task.Delay(50); // Đợi server sẵn sàng
                        using var warmupClient = new HttpClient();
                        warmupClient.Timeout = TimeSpan.FromSeconds(2);
                        
                        // Warm up nhiều lần để JIT compile tất cả code paths
                        await warmupClient.GetAsync($"http://localhost:{_config.Port}/status");
                        await warmupClient.GetAsync($"http://127.0.0.1:{_config.Port}/status");
                        await warmupClient.SendAsync(new HttpRequestMessage(HttpMethod.Options, $"http://localhost:{_config.Port}/"));
                        
                        Log("🔥 Warm up hoàn tất - Sẵn sàng phản hồi nhanh!");
                    }
                    catch { }
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

            // Thêm CORS headers cho TẤT CẢ response
            // Lấy Origin từ request (quan trọng cho Private Network Access)
            var origin = request.Headers.Get("Origin") ?? "*";
            response.AddHeader("Access-Control-Allow-Origin", origin);
            response.AddHeader("Access-Control-Allow-Methods", "GET, PUT, POST, DELETE, HEAD, OPTIONS, PROPFIND, LOCK, UNLOCK");
            response.AddHeader("Access-Control-Allow-Headers", "*");
            response.AddHeader("Access-Control-Allow-Credentials", "true");
            response.AddHeader("Access-Control-Allow-Private-Network", "true");
            response.AddHeader("Access-Control-Expose-Headers", "*");
            
            // WebDAV headers cho TẤT CẢ response - Word cần để nhận ra WebDAV
            response.AddHeader("DAV", "1, 2");
            response.AddHeader("MS-Author-Via", "DAV");
            response.AddHeader("Allow", "GET, PUT, HEAD, OPTIONS, PROPFIND, LOCK, UNLOCK");

            try
            {
                var path = request.Url?.AbsolutePath ?? "";
                var method = request.HttpMethod;

                Log($"📥 [{method}] {path}");

                // Xem log nhanh: http://localhost:{port}/logs
                if (path == "/logs")
                {
                    response.StatusCode = 200;
                    response.ContentType = "text/plain; charset=utf-8";
                    var text = string.Join(Environment.NewLine, _logBuffer);
                    var bufferLogs = Encoding.UTF8.GetBytes(text);
                    await response.OutputStream.WriteAsync(bufferLogs, 0, bufferLogs.Length);
                    return;
                }

                // Xử lý OPTIONS (preflight) cho TẤT CẢ paths
                if (method == "OPTIONS")
                {
                    HandleOptions(response, origin);
                    return;
                }

                // Xử lý PROPFIND cho root /files/ (Word có thể hỏi thư mục cha)
                if (path == "/files" || path == "/files/")
                {
                    if (method == "PROPFIND" || method == "OPTIONS")
                    {
                        HandleDirectoryPropfind(response);
                        return;
                    }
                }

                // Parse /files/:id/:userId/:object_type/:object_id/:edit_file_id/:filename
                if (path.StartsWith("/files/"))
                {
                    var parts = SplitWithLimit(path.Substring(7), '/', 6);
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
                                HandleOptions(response, origin);
                                break;
                            case "PROPFIND":
                                HandlePropfind(path, filename, response);
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

                // Quick status check - phản hồi nhanh nhất có thể
                if (path == "/status" || path == "/ping" || path == "/health")
                {
                    response.StatusCode = 200;
                    response.ContentType = "application/json";
                    var json = "{\"status\":\"ok\",\"running\":true}";
                    var buffer = Encoding.UTF8.GetBytes(json);
                    await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
                    return;
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
                            <p>API: {_fastestApiUrl ?? string.Join(", ", _config.GetApiUrls())}</p>
                        </body>
                        </html>";
                    var buffer = Encoding.UTF8.GetBytes(html);
                    response.ContentType = "text/html; charset=utf-8";
                    await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
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
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            var apiUrls = _config.GetApiUrls();
            
            Log($"   → Đang tải file từ server...");

            // Thử từng server với fallback
            HttpResponseMessage apiResponse = null;
            string usedServer = "";
            string lastError = "";
            
            // Ưu tiên server nhanh nhất, sau đó thử các server khác
            var orderedUrls = new System.Collections.Generic.List<string>();
            if (!string.IsNullOrEmpty(_fastestApiUrl))
            {
                orderedUrls.Add(_fastestApiUrl);
                foreach (var url in apiUrls)
                {
                    if (url != _fastestApiUrl) orderedUrls.Add(url);
                }
            }
            else
            {
                orderedUrls.AddRange(apiUrls);
            }
            
            foreach (var baseUrl in orderedUrls)
            {
                try
                {
                    var apiUrl = $"{baseUrl}{_config.ApiEndpoint}/{id}/{userId}/{objectType}/{objectId}/{editFileId}/{Uri.EscapeDataString(filename)}";
                    Log($"   → Thử server: {baseUrl}");
                    
                    apiResponse = await _httpClient.GetAsync(apiUrl);
                    
                    if (apiResponse != null && apiResponse.IsSuccessStatusCode)
                    {
                        usedServer = baseUrl;
                        // Cập nhật server nhanh nhất nếu khác
                        if (_fastestApiUrl != baseUrl)
                        {
                            _fastestApiUrl = baseUrl;
                            Log($"   → Cập nhật server nhanh nhất: {baseUrl}");
                        }
                        break;
                    }
                    else
                    {
                        lastError = $"Server {baseUrl} trả về lỗi {(int)(apiResponse?.StatusCode ?? 0)}";
                        Log($"   ⚠️ {lastError}");
                    }
                }
                catch (Exception ex)
                {
                    lastError = $"Không kết nối được {baseUrl}: {ex.Message}";
                    Log($"   ⚠️ {lastError}");
                }
            }
            
            var apiTime = stopwatch.ElapsedMilliseconds;
            
            if (apiResponse == null || !apiResponse.IsSuccessStatusCode)
            {
                string getError = apiResponse?.StatusCode switch
                {
                    System.Net.HttpStatusCode.NotFound => "File không tồn tại trên server",
                    System.Net.HttpStatusCode.Unauthorized => "Không có quyền truy cập. Vui lòng đăng nhập lại",
                    System.Net.HttpStatusCode.Forbidden => "Bị từ chối truy cập file này",
                    System.Net.HttpStatusCode.InternalServerError => "Lỗi server. Vui lòng thử lại sau",
                    _ => $"Không thể tải file từ bất kỳ server nào. {lastError}"
                };
                Log($"   ❌ LỖI MỞ FILE: {getError}");
                response.StatusCode = apiResponse != null ? (int)apiResponse.StatusCode : 503;
                return;
            }

            var content = await apiResponse.Content.ReadAsByteArrayAsync();
            var downloadTime = stopwatch.ElapsedMilliseconds;
            
            // Lưu filename gốc để đảm bảo PUT dùng đúng tên file
            var fileKey = $"{id}/{userId}/{objectType}/{objectId}/{editFileId}";
            _originalFilenames.AddOrUpdate(fileKey, filename, (key, oldValue) => filename);
            
            response.ContentType = GetMimeType(filename);
            // Encode filename cho UTF-8 (RFC 5987) - hỗ trợ tiếng Việt
            var safeFilename = Uri.EscapeDataString(filename);
            response.AddHeader("Content-Disposition", $"inline; filename=\"document{Path.GetExtension(filename)}\"; filename*=UTF-8''{safeFilename}");
            response.ContentLength64 = content.Length;
            await response.OutputStream.WriteAsync(content, 0, content.Length);
            
            stopwatch.Stop();
            var totalTime = stopwatch.ElapsedMilliseconds;
            var fileSizeKB = content.Length / 1024;
            
            Log($"   ✅ THÀNH CÔNG: {fileSizeKB}KB từ {usedServer} trong {totalTime}ms (API: {apiTime}ms)");
        }

        private async Task HandlePutAsync(string id, string userId, string objectType, string objectId, string editFileId, string filename, HttpListenerRequest request, HttpListenerResponse response)
        {
            // Read binary from Word
            using var ms = new MemoryStream();
            await request.InputStream.CopyToAsync(ms);
            var fileBytes = ms.ToArray();
            
            Log($"   📤 Nhận {fileBytes.Length} bytes từ Word");

            // Lấy filename gốc từ cache (đã lưu lúc GET) để đảm bảo không bị đổi tên
            var fileKey = $"{id}/{userId}/{objectType}/{objectId}/{editFileId}";
            string originalFilename = filename;
            if (_originalFilenames.TryGetValue(fileKey, out var cachedFilename))
            {
                originalFilename = cachedFilename;
                Log($"   📝 Dùng filename gốc: {originalFilename} (thay vì {filename})");
            }

            var apiUrls = _config.GetApiUrls();
            
            // Thử từng server với fallback
            HttpResponseMessage apiResponse = null;
            string usedServer = "";
            string lastError = "";
            
            // Ưu tiên server nhanh nhất, sau đó thử các server khác
            var orderedUrls = new System.Collections.Generic.List<string>();
            if (!string.IsNullOrEmpty(_fastestApiUrl))
            {
                orderedUrls.Add(_fastestApiUrl);
                foreach (var url in apiUrls)
                {
                    if (url != _fastestApiUrl) orderedUrls.Add(url);
                }
            }
            else
            {
                orderedUrls.AddRange(apiUrls);
            }
            
            foreach (var baseUrl in orderedUrls)
            {
                try
                {
                    var apiUrl = $"{baseUrl}{_config.ApiEndpoint}/{id}/{userId}/{objectType}/{objectId}/{editFileId}/{Uri.EscapeDataString(originalFilename)}";
                    Log($"   → Thử PUT lên: {baseUrl}");
                    Log($"   📋 Filename gửi lên BE: {originalFilename}");
                    Log($"   📋 URL API: {apiUrl}");

                    // Send as multipart/form-data (Company API expects this)
                    using var formContent = new MultipartFormDataContent();
                    var fileContent = new ByteArrayContent(fileBytes);
                    fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(GetMimeType(originalFilename));
                    
                    // Encode filename UTF-8 cho tiếng Việt - dùng filename gốc
                    var encodedFilename = Uri.EscapeDataString(originalFilename);
                    fileContent.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("form-data")
                    {
                        Name = "\"file\"",
                        FileName = $"\"{originalFilename}\"",
                        FileNameStar = encodedFilename
                    };
                    formContent.Add(fileContent);
                    
                    // Thêm filename vào DTO body để BE nhận được tên file đúng
                    // Thử nhiều field name phổ biến để tương thích với các DTO khác nhau
                    formContent.Add(new StringContent(originalFilename), "fileName");
                    formContent.Add(new StringContent(originalFilename), "file_name");
                    formContent.Add(new StringContent(originalFilename), "filename");
                    formContent.Add(new StringContent(originalFilename), "name");
                    
                    Log($"   📋 Content-Disposition: filename=\"{originalFilename}\", filename*={encodedFilename}");
                    Log($"   📋 Đã thêm filename vào DTO: {originalFilename}");

                    apiResponse = await _httpClient.PutAsync(apiUrl, formContent);
                    
                    if (apiResponse != null && apiResponse.IsSuccessStatusCode)
                    {
                        usedServer = baseUrl;
                        // Cập nhật server nhanh nhất nếu khác
                        if (_fastestApiUrl != baseUrl)
                        {
                            _fastestApiUrl = baseUrl;
                            Log($"   → Cập nhật server nhanh nhất: {baseUrl}");
                        }
                        break;
                    }
                    else
                    {
                        lastError = $"Server {baseUrl} trả về lỗi {(int)(apiResponse?.StatusCode ?? 0)}";
                        Log($"   ⚠️ {lastError}");
                    }
                }
                catch (Exception ex)
                {
                    lastError = $"Không kết nối được {baseUrl}: {ex.Message}";
                    Log($"   ⚠️ {lastError}");
                }
            }
            
            if (apiResponse == null || !apiResponse.IsSuccessStatusCode)
            {
                string putError = apiResponse?.StatusCode switch
                {
                    System.Net.HttpStatusCode.Unauthorized => "Không có quyền lưu. Vui lòng đăng nhập lại",
                    System.Net.HttpStatusCode.Forbidden => "Bị từ chối lưu file này",
                    System.Net.HttpStatusCode.RequestEntityTooLarge => "File quá lớn, không thể lưu",
                    System.Net.HttpStatusCode.InternalServerError => "Lỗi server. Vui lòng thử lại sau",
                    _ => $"Không thể lưu file lên bất kỳ server nào. {lastError}"
                };
                Log($"   ❌ LỖI LƯU FILE: {putError}");
                response.StatusCode = apiResponse != null ? (int)apiResponse.StatusCode : 503;
                return;
            }

            var result = await apiResponse.Content.ReadAsStringAsync();
            Log($"   ✅ LƯU THÀNH CÔNG lên {usedServer}!");

            response.ContentType = "application/json";
            var buffer = Encoding.UTF8.GetBytes(result);
            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
        }

        private async Task HandleHeadAsync(string id, string userId, string objectType, string objectId, string editFileId, string filename, HttpListenerResponse response)
        {
            var apiUrls = _config.GetApiUrls();
            
            // Ưu tiên server nhanh nhất, sau đó thử các server khác
            var orderedUrls = new System.Collections.Generic.List<string>();
            if (!string.IsNullOrEmpty(_fastestApiUrl))
            {
                orderedUrls.Add(_fastestApiUrl);
                foreach (var url in apiUrls)
                {
                    if (url != _fastestApiUrl) orderedUrls.Add(url);
                }
            }
            else
            {
                orderedUrls.AddRange(apiUrls);
            }
            
            foreach (var baseUrl in orderedUrls)
            {
                try
                {
                    var apiUrl = $"{baseUrl}{_config.ApiEndpoint}/{id}/{userId}/{objectType}/{objectId}/{editFileId}/{Uri.EscapeDataString(filename)}";
                    var request = new HttpRequestMessage(HttpMethod.Head, apiUrl);
                    var apiResponse = await _httpClient.SendAsync(request);
                    
                    if (apiResponse.IsSuccessStatusCode)
                    {
                        response.StatusCode = (int)apiResponse.StatusCode;
                        response.ContentType = GetMimeType(filename);
                        return;
                    }
                }
                catch { }
            }
            
            response.StatusCode = 404;
            response.ContentType = GetMimeType(filename);
        }

        private void HandleOptions(HttpListenerResponse response, string origin = "*")
        {
            response.AddHeader("Allow", "GET, PUT, HEAD, OPTIONS, PROPFIND, LOCK, UNLOCK");
            response.AddHeader("DAV", "1, 2");
            response.AddHeader("MS-Author-Via", "DAV");
            // CORS headers cho preflight request
            response.AddHeader("Access-Control-Allow-Origin", origin);
            response.AddHeader("Access-Control-Allow-Methods", "GET, PUT, POST, DELETE, HEAD, OPTIONS, PROPFIND, LOCK, UNLOCK");
            response.AddHeader("Access-Control-Allow-Headers", "*");
            response.AddHeader("Access-Control-Allow-Credentials", "true");
            response.AddHeader("Access-Control-Allow-Private-Network", "true");
            response.AddHeader("Access-Control-Expose-Headers", "*");
            response.AddHeader("Access-Control-Max-Age", "86400");
            // Một số bản Office/Word kén 204, trả 200 + Content-Length:0 để Word tiếp tục LOCK/PUT
            response.StatusCode = 200;
            response.ContentLength64 = 0;
        }

        private void HandlePropfind(string path, string filename, HttpListenerResponse response)
        {
            var mimeType = GetMimeType(filename);
            var etag = $"W/\"{path.GetHashCode():X}\"";
            var lastModified = DateTime.UtcNow.ToString("r");
            
            // Dùng path gốc từ request + đầy đủ properties cho Word
            var xml = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<D:multistatus xmlns:D=""DAV:"">
    <D:response>
        <D:href>{path}</D:href>
        <D:propstat>
            <D:prop>
                <D:displayname>{System.Security.SecurityElement.Escape(filename)}</D:displayname>
                <D:getcontenttype>{mimeType}</D:getcontenttype>
                <D:resourcetype/>
                <D:getetag>{etag}</D:getetag>
                <D:getlastmodified>{lastModified}</D:getlastmodified>
                <D:creationdate>{DateTime.UtcNow:yyyy-MM-ddTHH:mm:ssZ}</D:creationdate>
                <D:supportedlock>
                    <D:lockentry>
                        <D:lockscope><D:exclusive/></D:lockscope>
                        <D:locktype><D:write/></D:locktype>
                    </D:lockentry>
                    <D:lockentry>
                        <D:lockscope><D:shared/></D:lockscope>
                        <D:locktype><D:write/></D:locktype>
                    </D:lockentry>
                </D:supportedlock>
                <D:lockdiscovery/>
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
            var lockToken = $"opaquelocktoken:{Guid.NewGuid()}";
            var xml = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<D:prop xmlns:D=""DAV:"">
    <D:lockdiscovery>
        <D:activelock>
            <D:locktype><D:write/></D:locktype>
            <D:lockscope><D:exclusive/></D:lockscope>
            <D:depth>infinity</D:depth>
            <D:owner><D:href>WordWebDAV</D:href></D:owner>
            <D:timeout>Second-3600</D:timeout>
            <D:locktoken><D:href>{lockToken}</D:href></D:locktoken>
            <D:lockroot><D:href>/</D:href></D:lockroot>
        </D:activelock>
    </D:lockdiscovery>
</D:prop>";
            
            response.ContentType = "application/xml; charset=utf-8";
            response.StatusCode = 200;
            response.AddHeader("Lock-Token", $"<{lockToken}>");
            response.AddHeader("Timeout", "Second-3600");
            var buffer = Encoding.UTF8.GetBytes(xml);
            response.OutputStream.Write(buffer, 0, buffer.Length);
        }

        private void HandleUnlock(HttpListenerResponse response)
        {
            response.StatusCode = 204;
        }

        private void HandleDirectoryPropfind(HttpListenerResponse response)
        {
            var xml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<D:multistatus xmlns:D=""DAV:"">
    <D:response>
        <D:href>/files/</D:href>
        <D:propstat>
            <D:prop>
                <D:displayname>files</D:displayname>
                <D:resourcetype><D:collection/></D:resourcetype>
                <D:supportedlock>
                    <D:lockentry>
                        <D:lockscope><D:exclusive/></D:lockscope>
                        <D:locktype><D:write/></D:locktype>
                    </D:lockentry>
                </D:supportedlock>
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

        private void Log(string message)
        {
            var time = DateTime.Now.ToString("HH:mm:ss");
            var line = $"[{time}] {message}";

            // Đẩy vào queue log để xem qua /logs
            _logBuffer.Enqueue(line);
            while (_logBuffer.Count > MaxLogLines && _logBuffer.TryDequeue(out _)) { }

            if (OnLog != null)
                OnLog(line);
        }

        // Helper method cho .NET Framework (không có Split với limit)
        private static string[] SplitWithLimit(string input, char separator, int count)
        {
            var parts = input.Split(separator);
            if (parts.Length <= count) return parts;
            
            var result = new string[count];
            for (int i = 0; i < count - 1; i++)
                result[i] = parts[i];
            result[count - 1] = string.Join(separator.ToString(), parts, count - 1, parts.Length - count + 1);
            return result;
        }
    }
}
