// Local helper for the Zotero -> PowerPoint add-in.
//
// install.ps1 compiles this file with the C# compiler that ships with Windows
// (C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe), so the add-in needs
// no Python, no Node.js and no downloads at run time.
//
//   ZoteroHelper.exe [--api-port 8000] [--static-port 23000] [--www DIR]
//                    [--cert-dir DIR] [--no-static] [--upstream URL] [--log FILE] [--hidden]
//
// Two listeners in one process:
//   http://localhost:8000    JSON API used by the add-in and the Script Lab snippet
//   https://localhost:23000  the add-in web files (Office requires HTTPS for task panes)
//
// Written for the built-in compiler, which only accepts C# 5 syntax.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Net.Sockets;
using System.Security.Authentication;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Runtime.InteropServices;

class ZoteroHelper
{
    const int DefaultApiPort = 8000;
    const int DefaultStaticPort = 23000;
    const string DefaultUpstream = "http://127.0.0.1:23119";
    const string CertFileName = "localhost.pfx";
    const int MaxHeaderBytes = 64 * 1024;
    // The Zotero picker stays open until the user picks something, so give the
    // upstream call a generous ceiling instead of the 100 second default.
    const int UpstreamTimeoutMs = 10 * 60 * 1000;
    const string ZoteroDownError =
        "Could not connect to Zotero/BBT. Is Zotero running with the Better BibTeX plugin installed?";

    static readonly JavaScriptSerializer Serializer = new JavaScriptSerializer();
    static readonly object LogLock = new object();

    static StreamWriter logTarget;
    static string upstreamBase = DefaultUpstream;
    static string wwwRoot;
    static string certFile;
    static bool staticEnabled = true;
    static bool hideConsole;

    static void Main(string[] args)
    {
        int apiPort = DefaultApiPort;
        int staticPort = DefaultStaticPort;
        string wwwOption = null;
        string certDirOption = null;
        string logOption = null;

        for (int i = 0; i < args.Length; i++)
        {
            string name = args[i];
            if (name == "--help" || name == "-h")
            {
                PrintUsage();
                return;
            }
            if (name == "--no-static")
            {
                staticEnabled = false;
                continue;
            }
            if (name == "--hidden")
            {
                hideConsole = true;
                continue;
            }
            string value = i + 1 < args.Length ? args[i + 1] : null;
            if (value == null)
            {
                Console.Error.WriteLine("Missing value for " + name);
                Environment.Exit(2);
            }
            i++;
            if (name == "--api-port") apiPort = int.Parse(value, CultureInfo.InvariantCulture);
            else if (name == "--static-port") staticPort = int.Parse(value, CultureInfo.InvariantCulture);
            else if (name == "--www") wwwOption = value;
            else if (name == "--cert-dir") certDirOption = value;
            else if (name == "--upstream") upstreamBase = value.TrimEnd('/');
            else if (name == "--log") logOption = value;
            else
            {
                Console.Error.WriteLine("Unknown option: " + name);
                PrintUsage();
                Environment.Exit(2);
            }
        }

        if (logOption != null)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(logOption)));
            logTarget = new StreamWriter(logOption, true, new UTF8Encoding(false));
            logTarget.AutoFlush = true;
        }
        if (hideConsole) HideConsoleWindow();

        string exeDir = Path.GetDirectoryName(Path.GetFullPath(typeof(ZoteroHelper).Assembly.Location));
        wwwRoot = ResolveWww(exeDir, wwwOption);
        string certDir = certDirOption != null
            ? Path.GetFullPath(certDirOption)
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".office-addin-dev-certs");
        certFile = Path.Combine(certDir, CertFileName);

        // The upstream blocks while Zotero's picker is open, so make sure a slow
        // picker cannot starve the small requests that keep the pane responsive.
        int minWorkers; int minIo;
        ThreadPool.GetMinThreads(out minWorkers, out minIo);
        ThreadPool.SetMinThreads(Math.Max(minWorkers, 16), Math.Max(minIo, 16));

        StartApiServer(apiPort);

        if (!staticEnabled)
        {
            Log("Static file server disabled (--no-static)");
        }
        else if (wwwRoot == null)
        {
            Log("Add-in web files not found (use --www to point at the folder with taskpane.html)");
        }
        else if (!File.Exists(certFile))
        {
            Log("No certificate at " + certFile + " - run install.ps1 to create one");
        }
        else
        {
            StartStaticServer(staticPort);
        }

        new ManualResetEvent(false).WaitOne();
    }

    static void PrintUsage()
    {
        Console.WriteLine("ZoteroHelper.exe [--api-port 8000] [--static-port 23000] [--www DIR]");
        Console.WriteLine("                [--cert-dir DIR] [--no-static] [--upstream URL] [--log FILE] [--hidden]");
    }

    static string ResolveWww(string exeDir, string option)
    {
        if (option != null)
        {
            string candidate = Path.GetFullPath(option);
            return Directory.Exists(candidate) ? candidate : null;
        }
        string beside = Path.Combine(exeDir, "www");
        return Directory.Exists(beside) ? beside : null;
    }

    // ---------------------------------------------------------------- servers

    static void StartApiServer(int port)
    {
        TcpListener listener = Bind(port);
        Log("Local proxy server running on http://localhost:" + port.ToString(CultureInfo.InvariantCulture));
        Log("Forwarding /zotero and /bibliography to " + upstreamBase);
        Thread thread = new Thread(delegate() { AcceptLoop(listener, HandleApiClient); });
        thread.IsBackground = true;
        thread.Start();
    }

    static void StartStaticServer(int port)
    {
        TcpListener listener = Bind(port);
        Log("Add-in files served on https://localhost:" + port.ToString(CultureInfo.InvariantCulture)
            + " from " + wwwRoot);
        Thread thread = new Thread(delegate() { AcceptLoop(listener, HandleStaticClient); });
        thread.IsBackground = true;
        thread.Start();
    }

    static TcpListener Bind(int port)
    {
        TcpListener listener = new TcpListener(IPAddress.Loopback, port);
        try
        {
            listener.Start();
        }
        catch (SocketException e)
        {
            Log("FATAL: could not bind to port " + port.ToString(CultureInfo.InvariantCulture) + ": " + e.Message);
            Log("Hint: the helper may already be running. Check with: netstat -ano | findstr :" +
                port.ToString(CultureInfo.InvariantCulture));
            Environment.Exit(1);
        }
        return listener;
    }

    static void AcceptLoop(TcpListener listener, WaitCallback handler)
    {
        while (true)
        {
            TcpClient client;
            try
            {
                client = listener.AcceptTcpClient();
            }
            catch (Exception e)
            {
                Log("accept failed: " + e.Message);
                return;
            }
            ThreadPool.QueueUserWorkItem(handler, client);
        }
    }

    // ------------------------------------------------------------------- API

    static void HandleApiClient(object state)
    {
        TcpClient client = (TcpClient)state;
        try
        {
            using (NetworkStream stream = client.GetStream())
            {
                Request request = Request.Read(stream);
                if (request == null) return;

                if (request.Method == "OPTIONS")
                {
                    WriteResponse(stream, 200, "ok", "text/plain", new byte[0], true, false);
                }
                else if (request.Path == "/health")
                {
                    WriteJson(stream, 200, "{\"status\":\"ok\"}");
                }
                else if (request.Path == "/log" && request.Method == "POST")
                {
                    // The pane reports its own failures here, so they can be read from server.log.
                    string note = request.Body == null ? "" : request.Body.Trim();
                    if (note.Length > 4000) note = note.Substring(0, 4000);
                    Log("pane: " + note.Replace("\r", " ").Replace("\n", " "));
                    WriteJson(stream, 200, "{\"status\":\"logged\"}");
                }
                else if (request.Path == "/zotero")
                {
                    ProxyCitations(stream, request);
                }
                else if (request.Path == "/bibliography" && request.Method == "POST")
                {
                    GenerateBibliography(stream, request);
                }
                else
                {
                    WriteResponse(stream, 404, "Not Found", "text/plain", Encoding.UTF8.GetBytes("not found"), true, false);
                }
            }
        }
        catch (Exception e)
        {
            Log("api: " + e.Message);
        }
        finally
        {
            client.Close();
        }
    }

    static void ProxyCitations(Stream stream, Request request)
    {
        bool selected = request.Query != null && request.Query.IndexOf("selected=true", StringComparison.Ordinal) >= 0;
        string url = upstreamBase + "/better-bibtex/cayw?format=json" + (selected ? "&selected=true" : "");
        try
        {
            byte[] body = Fetch(url, "GET", null, "application/json");
            WriteResponse(stream, 200, "OK", "application/json", body, true, false);
        }
        catch (Exception e)
        {
            Log("error connecting to Zotero CAYW endpoint: " + e.Message);
            WriteJson(stream, 500, JsonError(ZoteroDownError));
        }
    }

    static void GenerateBibliography(Stream stream, Request request)
    {
        List<string> keys = new List<string>();
        string style = "apa";
        string contentType = "text";
        try
        {
            Dictionary<string, object> payload =
                (Dictionary<string, object>)Serializer.DeserializeObject(request.Body);
            // "html" gets the bibliography with real italics/bold/sub/superscript markup.
            object formatValue;
            if (payload.TryGetValue("format", out formatValue) && formatValue is string && (string)formatValue == "html")
            {
                contentType = "html";
            }
            object styleValue;
            if (payload.TryGetValue("style", out styleValue) && styleValue is string && ((string)styleValue).Length > 0)
            {
                style = (string)styleValue;
            }
            object keysValue;
            if (payload.TryGetValue("keys", out keysValue) && keysValue is object[])
            {
                object[] items = (object[])keysValue;
                for (int i = 0; i < items.Length; i++)
                {
                    if (items[i] is string) keys.Add((string)items[i]);
                }
            }
        }
        catch (Exception e)
        {
            Log("bibliography: could not read request: " + e.Message);
            WriteJson(stream, 500, JsonError("Internal server error: " + e.Message));
            return;
        }

        if (style == "apalike") style = "apa";
        if (keys.Count == 0)
        {
            WriteJson(stream, 400, JsonError("No citation keys provided"));
            return;
        }

        Dictionary<string, object> styleSpec = new Dictionary<string, object>();
        styleSpec.Add("id", style);
        styleSpec.Add("contentType", contentType);
        Dictionary<string, object> rpcCall = new Dictionary<string, object>();
        rpcCall.Add("jsonrpc", "2.0");
        rpcCall.Add("method", "item.bibliography");
        rpcCall.Add("params", new object[] { keys.ToArray(), styleSpec });

        try
        {
            byte[] requestBody = Encoding.UTF8.GetBytes(Serializer.Serialize(rpcCall));
            byte[] responseBody = Fetch(upstreamBase + "/better-bibtex/json-rpc", "POST", requestBody, "application/json");
            Dictionary<string, object> response =
                (Dictionary<string, object>)Serializer.DeserializeObject(Encoding.UTF8.GetString(responseBody));

            object error;
            if (response.TryGetValue("error", out error))
            {
                string message = DescribeError(error);
                Log("BBT JSON-RPC error: " + message);
                WriteJson(stream, 500, JsonError("Zotero/BBT Error: " + message));
                return;
            }

            object result;
            string bibliography = response.TryGetValue("result", out result) && result is string ? (string)result : "";
            Log("bibliography generated for " + keys.Count.ToString(CultureInfo.InvariantCulture)
                + " key(s), " + bibliography.Length.ToString(CultureInfo.InvariantCulture) + " characters");
            WriteJson(stream, 200, "{\"bibliography\":" + Serializer.Serialize(bibliography) + "}");
        }
        catch (Exception e)
        {
            Log("error connecting to BBT JSON-RPC endpoint: " + e.Message);
            WriteJson(stream, 500, JsonError(ZoteroDownError));
        }
    }

    static string DescribeError(object error)
    {
        Dictionary<string, object> errorObject = error as Dictionary<string, object>;
        if (errorObject != null)
        {
            object message;
            if (errorObject.TryGetValue("message", out message) && message is string) return (string)message;
        }
        return Serializer.Serialize(error);
    }

    static byte[] Fetch(string url, string method, byte[] body, string accept)
    {
        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
        request.Method = method;
        request.Accept = accept;
        request.Timeout = UpstreamTimeoutMs;
        request.ReadWriteTimeout = UpstreamTimeoutMs;
        request.KeepAlive = false;
        if (body != null)
        {
            request.ContentType = "application/json";
            request.ContentLength = body.Length;
            using (Stream requestStream = request.GetRequestStream())
            {
                requestStream.Write(body, 0, body.Length);
            }
        }
        using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
        using (MemoryStream buffer = new MemoryStream())
        {
            response.GetResponseStream().CopyTo(buffer);
            return buffer.ToArray();
        }
    }

    // ---------------------------------------------------------------- static

    static void HandleStaticClient(object state)
    {
        TcpClient client = (TcpClient)state;
        try
        {
            using (SslStream stream = new SslStream(client.GetStream(), false))
            {
                try
                {
                    X509Certificate2 certificate = new X509Certificate2(certFile, "");
                    stream.AuthenticateAsServer(certificate, false, SslProtocols.Tls12, false);
                }
                catch (Exception e)
                {
                    Log("static: TLS handshake failed: " + e.Message);
                    return;
                }

                Request request = Request.Read(stream);
                if (request == null) return;

                string path = request.Path;
                if (path.Length == 0 || path == "/") path = "/taskpane.html";
                ServeFile(stream, path);
            }
        }
        catch (Exception e)
        {
            Log("static: " + e.Message);
        }
        finally
        {
            client.Close();
        }
    }

    static void ServeFile(Stream stream, string requestPath)
    {
        string relative = requestPath.TrimStart('/');
        if (relative.IndexOf("..", StringComparison.Ordinal) >= 0)
        {
            WriteResponse(stream, 404, "Not Found", "text/plain", new byte[0], false, true);
            return;
        }

        string full;
        try
        {
            full = Path.GetFullPath(Path.Combine(wwwRoot, relative.Replace('/', Path.DirectorySeparatorChar)));
        }
        catch (Exception)
        {
            WriteResponse(stream, 404, "Not Found", "text/plain", new byte[0], false, true);
            return;
        }

        string rootWithSeparator = wwwRoot.EndsWith(Path.DirectorySeparatorChar.ToString())
            ? wwwRoot : wwwRoot + Path.DirectorySeparatorChar;
        if (!full.StartsWith(rootWithSeparator, StringComparison.OrdinalIgnoreCase) || !File.Exists(full))
        {
            Log("static: 404 " + requestPath);
            WriteResponse(stream, 404, "Not Found", "text/plain", new byte[0], false, true);
            return;
        }

        byte[] content = File.ReadAllBytes(full);
        WriteResponse(stream, 200, "OK", ContentTypeFor(full), content, false, true);
    }

    static string ContentTypeFor(string path)
    {
        string extension = Path.GetExtension(path).ToLowerInvariant();
        switch (extension)
        {
            case ".html": return "text/html; charset=utf-8";
            case ".js": return "application/javascript; charset=utf-8";
            case ".css": return "text/css; charset=utf-8";
            case ".json": return "application/json";
            case ".png": return "image/png";
            case ".jpg":
            case ".jpeg": return "image/jpeg";
            case ".svg": return "image/svg+xml";
            case ".ico": return "image/x-icon";
            default: return "application/octet-stream";
        }
    }

    // ------------------------------------------------------------ http plumbing

    class Request
    {
        public string Method = "";
        public string Path = "";
        public string Query = "";
        public string Body = "";

        public static Request Read(Stream stream)
        {
            byte[] head;
            byte[] leftover;
            if (!ReadHead(stream, out head, out leftover)) return null;

            string text = Encoding.UTF8.GetString(head);
            string[] lines = text.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
            if (lines.Length == 0 || lines[0].Length == 0) return null;

            Request request = new Request();
            string[] requestLine = lines[0].Split(' ');
            if (requestLine.Length < 2) return null;
            request.Method = requestLine[0].ToUpperInvariant();
            string target = requestLine[1];
            int questionMark = target.IndexOf('?');
            request.Path = Uri.UnescapeDataString(questionMark >= 0 ? target.Substring(0, questionMark) : target);
            request.Query = questionMark >= 0 ? target.Substring(questionMark + 1) : "";

            int contentLength = 0;
            for (int i = 1; i < lines.Length; i++)
            {
                int colon = lines[i].IndexOf(':');
                if (colon <= 0) continue;
                string name = lines[i].Substring(0, colon).Trim();
                string value = lines[i].Substring(colon + 1).Trim();
                if (string.Equals(name, "Content-Length", StringComparison.OrdinalIgnoreCase))
                {
                    int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out contentLength);
                }
            }

            byte[] body = ReadBody(stream, leftover, contentLength);
            request.Body = Encoding.UTF8.GetString(body);
            return request;
        }

        static bool ReadHead(Stream stream, out byte[] head, out byte[] leftover)
        {
            head = null;
            leftover = null;
            List<byte> buffer = new List<byte>();
            byte[] chunk = new byte[4096];
            while (true)
            {
                int read;
                try
                {
                    read = stream.Read(chunk, 0, chunk.Length);
                }
                catch (IOException)
                {
                    return false;
                }
                if (read <= 0) return false;
                for (int i = 0; i < read; i++) buffer.Add(chunk[i]);

                int terminatorLength;
                int end = FindHeaderEnd(buffer, out terminatorLength);
                if (end >= 0)
                {
                    head = buffer.GetRange(0, end).ToArray();
                    int bodyStart = end + terminatorLength;
                    leftover = buffer.GetRange(bodyStart, buffer.Count - bodyStart).ToArray();
                    return true;
                }
                if (buffer.Count > MaxHeaderBytes) return false;
            }
        }

        static int FindHeaderEnd(List<byte> buffer, out int terminatorLength)
        {
            terminatorLength = 0;
            for (int i = 3; i < buffer.Count; i++)
            {
                if (buffer[i - 3] == 13 && buffer[i - 2] == 10 && buffer[i - 1] == 13 && buffer[i] == 10)
                {
                    terminatorLength = 4;
                    return i - 3;
                }
            }
            for (int i = 1; i < buffer.Count; i++)
            {
                if (buffer[i - 1] == 10 && buffer[i] == 10)
                {
                    terminatorLength = 2;
                    return i - 1;
                }
            }
            return -1;
        }

        static byte[] ReadBody(Stream stream, byte[] leftover, int contentLength)
        {
            if (contentLength <= 0) return new byte[0];
            MemoryStream body = new MemoryStream();
            int take = Math.Min(contentLength, leftover.Length);
            body.Write(leftover, 0, take);
            byte[] chunk = new byte[8192];
            while (body.Length < contentLength)
            {
                int read;
                try
                {
                    read = stream.Read(chunk, 0, (int)Math.Min(chunk.Length, contentLength - body.Length));
                }
                catch (IOException)
                {
                    break;
                }
                if (read <= 0) break;
                body.Write(chunk, 0, read);
            }
            return body.ToArray();
        }
    }

    static void WriteJson(Stream stream, int status, string json)
    {
        WriteResponse(stream, status, status == 200 ? "OK" : "Error", "application/json",
            Encoding.UTF8.GetBytes(json), true, false);
    }

    static string JsonError(string message)
    {
        return "{\"error\":" + Serializer.Serialize(message) + "}";
    }

    static void WriteResponse(Stream stream, int status, string reason, string contentType,
        byte[] body, bool cors, bool noStore)
    {
        StringBuilder head = new StringBuilder();
        head.Append("HTTP/1.1 ").Append(status.ToString(CultureInfo.InvariantCulture)).Append(' ').Append(reason).Append("\r\n");
        head.Append("Content-Type: ").Append(contentType).Append("\r\n");
        head.Append("Content-Length: ").Append(body.Length.ToString(CultureInfo.InvariantCulture)).Append("\r\n");
        head.Append("Connection: close\r\n");
        if (noStore) head.Append("Cache-Control: no-store\r\n");
        if (cors)
        {
            head.Append("Access-Control-Allow-Origin: *\r\n");
            head.Append("Access-Control-Allow-Methods: GET, POST, OPTIONS\r\n");
            // The pane's JSON POST is preflighted, so Chromium needs this header to let it through.
            head.Append("Access-Control-Allow-Headers: X-Requested-With, Content-Type\r\n");
        }
        head.Append("\r\n");

        byte[] headBytes = Encoding.ASCII.GetBytes(head.ToString());
        try
        {
            stream.Write(headBytes, 0, headBytes.Length);
            if (body.Length > 0) stream.Write(body, 0, body.Length);
            stream.Flush();
        }
        catch (IOException)
        {
            // The client went away (closed pane, cancelled request); nothing to do.
        }
        catch (ObjectDisposedException)
        {
        }
    }

    [DllImport("kernel32.dll")]
    static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    static extern bool ShowWindow(IntPtr window, int command);

    static void HideConsoleWindow()
    {
        IntPtr console = GetConsoleWindow();
        if (console != IntPtr.Zero) ShowWindow(console, 0);   // 0 = SW_HIDE
    }

    static void Log(string message)
    {
        string line = "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + "] " + message;
        lock (LogLock)
        {
            if (logTarget != null) logTarget.WriteLine(line);
            else Console.Out.WriteLine(line);
        }
    }
}
