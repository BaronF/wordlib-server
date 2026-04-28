import com.sun.net.httpserver.*;
import java.io.*;
import java.net.*;
import java.sql.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;
import java.text.SimpleDateFormat;

public class WordLibServer {

    static final int PORT = 8080;
    static String SCRIPT_DIR;
    static String DB_FILE;
    static String LOG_FILE;
    static Connection globalConn;
    static final Object DB_LOCK = new Object();

    public static void main(String[] args) throws Exception {
        System.setOut(new PrintStream(System.out, true, "UTF-8"));
        SCRIPT_DIR = new File(WordLibServer.class.getProtectionDomain().getCodeSource().getLocation().toURI()).getParent();
        if (SCRIPT_DIR == null) SCRIPT_DIR = System.getProperty("user.dir");
        DB_FILE = SCRIPT_DIR + File.separator + "wordlib.db";
        LOG_FILE = SCRIPT_DIR + File.separator + "server.log";
        System.out.println("========================================");
        System.out.println("  \u8BCD\u5E93\u7BA1\u7406\u7CFB\u7EDF v3 (Java)");
        System.out.println("========================================");
        System.out.println("\u5DE5\u4F5C\u76EE\u5F55: " + SCRIPT_DIR);
        System.out.println("\u6570\u636E\u5E93: " + DB_FILE);
        System.out.println("\u65E5\u5FD7\u6587\u4EF6: " + LOG_FILE);
        System.out.flush();
        initDatabase();
        System.out.println("\u6570\u636E\u5E93\u5C31\u7EEA");
        System.out.flush();
        writeLog("\u670D\u52A1\u5DF2\u542F\u52A8");
        HttpServer server = HttpServer.create(new InetSocketAddress(PORT), 0);
        server.createContext("/api/words", new WordsHandler());
        server.createContext("/api/roots", new RootsHandler());
        server.createContext("/api/stats", new StatsHandler());
        server.createContext("/api/check_init", new CheckInitHandler());
        server.createContext("/api/init_words", new InitWordsHandler());
        server.createContext("/api/init_roots", new InitRootsHandler());
        server.createContext("/api/import_check", new ImportCheckHandler());
        server.createContext("/", new StaticHandler());
        server.setExecutor(java.util.concurrent.Executors.newFixedThreadPool(4));
        server.start();
        System.out.println("\u5DF2\u542F\u52A8: http://localhost:" + PORT);
        System.out.println("\u6309 Ctrl+C \u505C\u6B62\u670D\u52A1");
        System.out.flush();
    }

    static synchronized void writeLog(String msg) {
        String ts = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new java.util.Date());
        String line = "[" + ts + "] " + msg;
        System.out.println(line);
        System.out.flush();
        try (FileOutputStream fos = new FileOutputStream(LOG_FILE, true);
             OutputStreamWriter w = new OutputStreamWriter(fos, StandardCharsets.UTF_8)) {
            w.write(line + "\n");
        } catch (Exception e) { }
    }

    static synchronized Connection getDb() throws SQLException {
        if (globalConn == null || globalConn.isClosed()) {
            globalConn = DriverManager.getConnection("jdbc:sqlite:" + DB_FILE);
            globalConn.createStatement().execute("PRAGMA journal_mode=WAL");
            globalConn.createStatement().execute("PRAGMA busy_timeout=5000");
        }
        return globalConn;
    }

    static void initDatabase() throws Exception {
        Class.forName("org.sqlite.JDBC");
        Connection conn = getDb();
        synchronized (DB_LOCK) {
            Statement stmt = conn.createStatement();
            stmt.executeUpdate(
                "CREATE TABLE IF NOT EXISTS words (" +
                "id INTEGER PRIMARY KEY AUTOINCREMENT," +
                "cn TEXT NOT NULL DEFAULT ''," +
                "en TEXT NOT NULL DEFAULT ''," +
                "cat TEXT DEFAULT ''," +
                "roots TEXT DEFAULT ''," +
                "score REAL DEFAULT 0," +
                "abbr TEXT DEFAULT ''," +
                "cnDesc TEXT DEFAULT ''," +
                "enDesc TEXT DEFAULT ''," +
                "ref TEXT DEFAULT ''," +
                "status TEXT DEFAULT 'draft'," +
                "time TEXT DEFAULT (date('now','localtime')))"
            );
            stmt.executeUpdate(
                "CREATE TABLE IF NOT EXISTS roots (" +
                "id INTEGER PRIMARY KEY AUTOINCREMENT," +
                "name TEXT NOT NULL DEFAULT ''," +
                "en TEXT DEFAULT ''," +
                "mean TEXT DEFAULT ''," +
                "src TEXT DEFAULT ''," +
                "cat TEXT DEFAULT ''," +
                "status TEXT DEFAULT 'draft'," +
                "examples TEXT DEFAULT '[]')"
            );
            stmt.close();
        }
    }

    // ===== JSON Utils =====
    static String readBody(HttpExchange ex) throws IOException {
        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        byte[] tmp = new byte[4096]; int n;
        InputStream is = ex.getRequestBody();
        while ((n = is.read(tmp)) != -1) buf.write(tmp, 0, n);
        return new String(buf.toByteArray(), StandardCharsets.UTF_8);
    }

    static void sendJson(HttpExchange ex, int code, String json) throws IOException {
        byte[] bytes = json.getBytes(StandardCharsets.UTF_8);
        ex.getResponseHeaders().set("Content-Type", "application/json; charset=utf-8");
        addCors(ex);
        ex.sendResponseHeaders(code, bytes.length);
        ex.getResponseBody().write(bytes);
        ex.getResponseBody().close();
    }

    static void addCors(HttpExchange ex) {
        ex.getResponseHeaders().set("Access-Control-Allow-Origin", "*");
        ex.getResponseHeaders().set("Access-Control-Allow-Methods", "GET,POST,PUT,DELETE,OPTIONS");
        ex.getResponseHeaders().set("Access-Control-Allow-Headers", "Content-Type");
    }

    static boolean handleOptions(HttpExchange ex) throws IOException {
        if ("OPTIONS".equalsIgnoreCase(ex.getRequestMethod())) {
            addCors(ex); ex.sendResponseHeaders(204, -1); ex.close();
            writeLog(reqLine(ex));
            return true;
        }
        return false;
    }

    static String esc(String s) {
        if (s == null) return "";
        return s.replace("\\", "\\\\").replace("\"", "\\\"").replace("\n", "\\n").replace("\r", "\\r").replace("\t", "\\t");
    }

    static String jsonStr(String json, String key) {
        String search = "\"" + key + "\"";
        int idx = json.indexOf(search);
        if (idx < 0) return "";
        idx = json.indexOf(":", idx + search.length());
        if (idx < 0) return "";
        idx++;
        while (idx < json.length() && (json.charAt(idx) == ' ' || json.charAt(idx) == '\t')) idx++;
        if (idx >= json.length()) return "";
        if (json.charAt(idx) == '"') {
            idx++;
            StringBuilder sb = new StringBuilder();
            while (idx < json.length() && json.charAt(idx) != '"') {
                if (json.charAt(idx) == '\\' && idx + 1 < json.length()) {
                    char next = json.charAt(idx + 1);
                    if (next == '"') { sb.append('"'); idx += 2; }
                    else if (next == '\\') { sb.append('\\'); idx += 2; }
                    else if (next == 'n') { sb.append('\n'); idx += 2; }
                    else if (next == 'r') { sb.append('\r'); idx += 2; }
                    else if (next == 't') { sb.append('\t'); idx += 2; }
                    else { sb.append(json.charAt(idx)); idx++; }
                } else { sb.append(json.charAt(idx)); idx++; }
            }
            return sb.toString();
        }
        int start = idx;
        while (idx < json.length() && json.charAt(idx) != ',' && json.charAt(idx) != '}' && json.charAt(idx) != ']') idx++;
        return json.substring(start, idx).trim();
    }

    static double jsonNum(String json, String key) {
        String v = jsonStr(json, key);
        if (v.isEmpty()) return 0;
        try { return Double.parseDouble(v); } catch (Exception e) { return 0; }
    }

    static List<String> jsonArray(String json) {
        List<String> list = new ArrayList<>();
        json = json.trim();
        if (!json.startsWith("[")) return list;
        json = json.substring(1, json.length() - 1).trim();
        if (json.isEmpty()) return list;
        int depth = 0; int start = 0;
        for (int i = 0; i < json.length(); i++) {
            char c = json.charAt(i);
            if (c == '{' || c == '[') depth++;
            else if (c == '}' || c == ']') depth--;
            else if (c == '"') { i++; while (i < json.length() && json.charAt(i) != '"') { if (json.charAt(i) == '\\') i++; i++; } }
            else if (c == ',' && depth == 0) { list.add(json.substring(start, i).trim()); start = i + 1; }
        }
        if (start < json.length()) list.add(json.substring(start).trim());
        return list;
    }

    static String wordToJson(ResultSet rs) throws SQLException {
        return "{\"id\":" + rs.getInt("id") +
            ",\"cn\":\"" + esc(rs.getString("cn")) + "\"" +
            ",\"en\":\"" + esc(rs.getString("en")) + "\"" +
            ",\"cat\":\"" + esc(rs.getString("cat")) + "\"" +
            ",\"roots\":\"" + esc(rs.getString("roots")) + "\"" +
            ",\"score\":" + rs.getDouble("score") +
            ",\"abbr\":\"" + esc(rs.getString("abbr")) + "\"" +
            ",\"cnDesc\":\"" + esc(rs.getString("cnDesc")) + "\"" +
            ",\"enDesc\":\"" + esc(rs.getString("enDesc")) + "\"" +
            ",\"ref\":\"" + esc(rs.getString("ref")) + "\"" +
            ",\"status\":\"" + esc(rs.getString("status")) + "\"" +
            ",\"time\":\"" + esc(rs.getString("time")) + "\"}";
    }

    static String rootToJson(ResultSet rs) throws SQLException {
        return "{\"id\":" + rs.getInt("id") +
            ",\"name\":\"" + esc(rs.getString("name")) + "\"" +
            ",\"en\":\"" + esc(rs.getString("en")) + "\"" +
            ",\"mean\":\"" + esc(rs.getString("mean")) + "\"" +
            ",\"src\":\"" + esc(rs.getString("src")) + "\"" +
            ",\"cat\":\"" + esc(rs.getString("cat")) + "\"" +
            ",\"status\":\"" + esc(rs.getString("status")) + "\"" +
            ",\"examples\":\"" + esc(rs.getString("examples")) + "\"}";
    }

    static Map<String, String> parseQuery(String query) {
        Map<String, String> map = new LinkedHashMap<>();
        if (query == null || query.isEmpty()) return map;
        for (String p : query.split("&")) {
            String[] kv = p.split("=", 2);
            try { map.put(kv[0], kv.length > 1 ? URLDecoder.decode(kv[1], "UTF-8") : ""); }
            catch (Exception e) { map.put(kv[0], kv.length > 1 ? kv[1] : ""); }
        }
        return map;
    }

    static double calcSimilarity(String s1, String s2) {
        if (s1 == null || s2 == null) return 0;
        s1 = s1.toLowerCase().replaceAll("[\\[\\]\"\\s]", "");
        s2 = s2.toLowerCase().replaceAll("[\\[\\]\"\\s]", "");
        if (s1.equals(s2)) return 1.0;
        int len1 = s1.length(), len2 = s2.length();
        if (len1 == 0 || len2 == 0) return 0;
        if (len1 >= 2 && len2 >= 2) {
            if (s2.contains(s1)) return Math.max(0.7, (double) len1 / len2);
            if (s1.contains(s2)) return Math.max(0.7, (double) len2 / len1);
        }
        int[][] dp = new int[len1 + 1][len2 + 1];
        for (int i = 0; i <= len1; i++) dp[i][0] = i;
        for (int j = 0; j <= len2; j++) dp[0][j] = j;
        for (int i = 1; i <= len1; i++)
            for (int j = 1; j <= len2; j++) {
                int cost = s1.charAt(i - 1) == s2.charAt(j - 1) ? 0 : 1;
                dp[i][j] = Math.min(Math.min(dp[i-1][j]+1, dp[i][j-1]+1), dp[i-1][j-1]+cost);
            }
        return 1.0 - (double) dp[len1][len2] / Math.max(len1, len2);
    }

    static String reqLine(HttpExchange ex) {
        String q = ex.getRequestURI().getQuery();
        String uri = ex.getRequestURI().getPath() + (q != null && !q.isEmpty() ? "?" + q : "");
        return ex.getRequestMethod().toUpperCase() + " " + uri + " " + ex.getProtocol();
    }

    // ===== WordsHandler =====
    static class WordsHandler implements HttpHandler {
        public void handle(HttpExchange ex) throws IOException {
            if (handleOptions(ex)) return;
            String method = ex.getRequestMethod().toUpperCase();
            String path = ex.getRequestURI().getPath();
            try {
                Connection conn = getDb();
                synchronized (DB_LOCK) {
                    if ("GET".equals(method)) {
                        Map<String, String> q = parseQuery(ex.getRequestURI().getQuery());
                        String search = q.getOrDefault("search", "");
                        String cat = q.getOrDefault("cat", "");
                        String status = q.getOrDefault("status", "");
                        int page = 1, size = 50;
                        try { page = Integer.parseInt(q.getOrDefault("page", "1")); } catch (Exception e) {}
                        try { size = Integer.parseInt(q.getOrDefault("size", "50")); } catch (Exception e) {}
                        StringBuilder where = new StringBuilder();
                        List<String> args = new ArrayList<>();
                        if (!search.isEmpty()) {
                            where.append("(cn LIKE ? OR en LIKE ? OR roots LIKE ? OR cnDesc LIKE ?)");
                            for (int i = 0; i < 4; i++) args.add("%" + search + "%");
                        }
                        if (!cat.isEmpty()) { if (where.length() > 0) where.append(" AND "); where.append("cat=?"); args.add(cat); }
                        if (!status.isEmpty()) { if (where.length() > 0) where.append(" AND "); where.append("status=?"); args.add(status); }
                        String w = where.length() > 0 ? " WHERE " + where : "";
                        PreparedStatement ps1 = conn.prepareStatement("SELECT COUNT(*) FROM words" + w);
                        for (int i = 0; i < args.size(); i++) ps1.setString(i + 1, args.get(i));
                        ResultSet rs1 = ps1.executeQuery(); int total = rs1.next() ? rs1.getInt(1) : 0; rs1.close(); ps1.close();
                        PreparedStatement ps2 = conn.prepareStatement("SELECT * FROM words" + w + " ORDER BY id DESC LIMIT ? OFFSET ?");
                        for (int i = 0; i < args.size(); i++) ps2.setString(i + 1, args.get(i));
                        ps2.setInt(args.size() + 1, size); ps2.setInt(args.size() + 2, (page - 1) * size);
                        ResultSet rs2 = ps2.executeQuery();
                        StringBuilder sb = new StringBuilder("["); boolean first = true;
                        while (rs2.next()) { if (!first) sb.append(","); sb.append(wordToJson(rs2)); first = false; }
                        sb.append("]"); rs2.close(); ps2.close();
                        writeLog("\u2705 \u67E5\u8BE2\u8BCD\u6761 - \u5171" + total + "\u6761\uFF0C\u7B2C" + page + "\u9875\uFF0C\u6BCF\u9875" + size + "\u6761");
                        writeLog(reqLine(ex));
                        sendJson(ex, 200, "{\"total\":" + total + ",\"page\":" + page + ",\"size\":" + size + ",\"data\":" + sb + "}");
                    } else if ("POST".equals(method)) {
                        String body = readBody(ex);
                        String cn = jsonStr(body, "cn"), en = jsonStr(body, "en"), catV = jsonStr(body, "cat");
                        String roots = jsonStr(body, "roots"), abbr = jsonStr(body, "abbr");
                        String cnDesc = jsonStr(body, "cnDesc"), enDesc = jsonStr(body, "enDesc");
                        String ref = jsonStr(body, "ref"), statusV = jsonStr(body, "status"), time = jsonStr(body, "time");
                        double score = jsonNum(body, "score");
                        PreparedStatement ps = conn.prepareStatement(
                            "INSERT INTO words(cn,en,cat,roots,score,abbr,cnDesc,enDesc,ref,status,time) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                            Statement.RETURN_GENERATED_KEYS);
                        ps.setString(1, cn); ps.setString(2, en); ps.setString(3, catV);
                        ps.setString(4, roots); ps.setDouble(5, score); ps.setString(6, abbr);
                        ps.setString(7, cnDesc); ps.setString(8, enDesc); ps.setString(9, ref);
                        ps.setString(10, statusV.isEmpty() ? "draft" : statusV);
                        ps.setString(11, time.isEmpty() ? new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) : time);
                        ps.executeUpdate();
                        ResultSet keys = ps.getGeneratedKeys(); int newId = keys.next() ? keys.getInt(1) : 0; keys.close(); ps.close();
                        writeLog("\u2705 \u65B0\u589E\u8BCD\u6761 - ID:" + newId + ", " + cn + "(" + en + "), \u5206\u7C7B:" + catV + ", \u72B6\u6001:" + (statusV.isEmpty()?"draft":statusV));
                        writeLog(reqLine(ex));
                        sendJson(ex, 201, "{\"id\":" + newId + ",\"msg\":\"ok\"}");
                    } else if ("PUT".equals(method)) {
                        int wid = Integer.parseInt(path.substring(path.lastIndexOf('/') + 1));
                        String body = readBody(ex);
                        PreparedStatement ps = conn.prepareStatement(
                            "UPDATE words SET cn=?,en=?,cat=?,roots=?,score=?,abbr=?,cnDesc=?,enDesc=?,ref=?,status=?,time=? WHERE id=?");
                        ps.setString(1, jsonStr(body,"cn")); ps.setString(2, jsonStr(body,"en")); ps.setString(3, jsonStr(body,"cat"));
                        ps.setString(4, jsonStr(body,"roots")); ps.setDouble(5, jsonNum(body,"score")); ps.setString(6, jsonStr(body,"abbr"));
                        ps.setString(7, jsonStr(body,"cnDesc")); ps.setString(8, jsonStr(body,"enDesc")); ps.setString(9, jsonStr(body,"ref"));
                        ps.setString(10, jsonStr(body,"status")); ps.setString(11, jsonStr(body,"time")); ps.setInt(12, wid);
                        ps.executeUpdate(); ps.close();
                        writeLog("\u2705 \u66F4\u65B0\u8BCD\u6761 - ID:" + wid + ", " + jsonStr(body,"cn") + "(" + jsonStr(body,"en") + "), \u72B6\u6001:" + jsonStr(body,"status"));
                        writeLog(reqLine(ex));
                        sendJson(ex, 200, "{\"msg\":\"ok\"}");
                    } else if ("DELETE".equals(method)) {
                        int wid = Integer.parseInt(path.substring(path.lastIndexOf('/') + 1));
                        PreparedStatement ps = conn.prepareStatement("DELETE FROM words WHERE id=?");
                        ps.setInt(1, wid); ps.executeUpdate(); ps.close();
                        writeLog("\u2705 \u5220\u9664\u8BCD\u6761 - ID:" + wid);
                        writeLog(reqLine(ex));
                        sendJson(ex, 200, "{\"msg\":\"ok\"}");
                    }
                }
            } catch (Exception e) {
                writeLog("\u274C \u8BCD\u6761\u64CD\u4F5C\u5F02\u5E38: " + e.getMessage());
                sendJson(ex, 500, "{\"error\":\"" + esc(e.getMessage()) + "\"}");
            }
        }
    }

    // ===== RootsHandler =====
    static class RootsHandler implements HttpHandler {
        public void handle(HttpExchange ex) throws IOException {
            if (handleOptions(ex)) return;
            String method = ex.getRequestMethod().toUpperCase();
            String path = ex.getRequestURI().getPath();
            try {
                Connection conn = getDb();
                synchronized (DB_LOCK) {
                    if ("GET".equals(method)) {
                        Map<String, String> q = parseQuery(ex.getRequestURI().getQuery());
                        String search = q.getOrDefault("search", ""), cat = q.getOrDefault("cat", "");
                        StringBuilder where = new StringBuilder(); List<String> args = new ArrayList<>();
                        if (!search.isEmpty()) { where.append("(name LIKE ? OR en LIKE ? OR mean LIKE ?)"); for (int i=0;i<3;i++) args.add("%"+search+"%"); }
                        if (!cat.isEmpty()) { if (where.length()>0) where.append(" AND "); where.append("cat=?"); args.add(cat); }
                        String w = where.length() > 0 ? " WHERE " + where : "";
                        PreparedStatement ps = conn.prepareStatement("SELECT * FROM roots" + w + " ORDER BY id DESC");
                        for (int i = 0; i < args.size(); i++) ps.setString(i+1, args.get(i));
                        ResultSet rs = ps.executeQuery();
                        StringBuilder sb = new StringBuilder("["); boolean first = true; int count = 0;
                        while (rs.next()) { if (!first) sb.append(","); sb.append(rootToJson(rs)); first = false; count++; }
                        sb.append("]"); rs.close(); ps.close();
                        writeLog("\u2705 \u67E5\u8BE2\u8BCD\u6839 - \u5171" + count + "\u6761");
                        writeLog(reqLine(ex));
                        sendJson(ex, 200, sb.toString());
                    } else if ("POST".equals(method)) {
                        String body = readBody(ex);
                        String name = jsonStr(body,"name"), en = jsonStr(body,"en"), mean = jsonStr(body,"mean");
                        String src = jsonStr(body,"src"), catV = jsonStr(body,"cat"), statusV = jsonStr(body,"status");
                        String examples = jsonStr(body,"examples"); if (examples.isEmpty()) examples = "[]";
                        PreparedStatement ps = conn.prepareStatement("INSERT INTO roots(name,en,mean,src,cat,status,examples) VALUES(?,?,?,?,?,?,?)", Statement.RETURN_GENERATED_KEYS);
                        ps.setString(1,name); ps.setString(2,en); ps.setString(3,mean); ps.setString(4,src); ps.setString(5,catV);
                        ps.setString(6, statusV.isEmpty()?"draft":statusV); ps.setString(7,examples);
                        ps.executeUpdate(); ResultSet keys = ps.getGeneratedKeys(); int newId = keys.next()?keys.getInt(1):0; keys.close(); ps.close();
                        writeLog("\u2705 \u65B0\u589E\u8BCD\u6839 - ID:" + newId + ", " + name + "(" + en + "), \u72B6\u6001:" + (statusV.isEmpty()?"draft":statusV));
                        writeLog(reqLine(ex));
                        sendJson(ex, 201, "{\"id\":" + newId + ",\"msg\":\"ok\"}");
                    } else if ("PUT".equals(method)) {
                        int rid = Integer.parseInt(path.substring(path.lastIndexOf('/')+1));
                        String body = readBody(ex);
                        String examples = jsonStr(body,"examples"); if (examples.isEmpty()) examples = "[]";
                        PreparedStatement ps = conn.prepareStatement("UPDATE roots SET name=?,en=?,mean=?,src=?,cat=?,status=?,examples=? WHERE id=?");
                        ps.setString(1,jsonStr(body,"name")); ps.setString(2,jsonStr(body,"en")); ps.setString(3,jsonStr(body,"mean"));
                        ps.setString(4,jsonStr(body,"src")); ps.setString(5,jsonStr(body,"cat")); ps.setString(6,jsonStr(body,"status"));
                        ps.setString(7,examples); ps.setInt(8,rid);
                        ps.executeUpdate(); ps.close();
                        writeLog("\u2705 \u66F4\u65B0\u8BCD\u6839 - ID:" + rid + ", " + jsonStr(body,"name") + "(" + jsonStr(body,"en") + "), \u72B6\u6001:" + jsonStr(body,"status"));
                        writeLog(reqLine(ex));
                        sendJson(ex, 200, "{\"msg\":\"ok\"}");
                    } else if ("DELETE".equals(method)) {
                        int rid = Integer.parseInt(path.substring(path.lastIndexOf('/')+1));
                        PreparedStatement ps = conn.prepareStatement("DELETE FROM roots WHERE id=?");
                        ps.setInt(1,rid); ps.executeUpdate(); ps.close();
                        writeLog("\u2705 \u5220\u9664\u8BCD\u6839 - ID:" + rid);
                        writeLog(reqLine(ex));
                        sendJson(ex, 200, "{\"msg\":\"ok\"}");
                    }
                }
            } catch (Exception e) {
                writeLog("\u274C \u8BCD\u6839\u64CD\u4F5C\u5F02\u5E38: " + e.getMessage());
                sendJson(ex, 500, "{\"error\":\"" + esc(e.getMessage()) + "\"}");
            }
        }
    }

    // ===== StatsHandler =====
    static class StatsHandler implements HttpHandler {
        public void handle(HttpExchange ex) throws IOException {
            if (handleOptions(ex)) return;
            try {
                Connection conn = getDb();
                synchronized (DB_LOCK) {
                    int wc=0, rc=0;
                    ResultSet r1 = conn.createStatement().executeQuery("SELECT COUNT(*) FROM words");
                    if (r1.next()) wc = r1.getInt(1); r1.close();
                    ResultSet r2 = conn.createStatement().executeQuery("SELECT COUNT(*) FROM roots");
                    if (r2.next()) rc = r2.getInt(1); r2.close();
                    sendJson(ex, 200, "{\"wordCount\":" + wc + ",\"rootCount\":" + rc + "}");
                }
            } catch (Exception e) { sendJson(ex, 500, "{\"error\":\"" + esc(e.getMessage()) + "\"}"); }
        }
    }

    // ===== CheckInitHandler =====
    static class CheckInitHandler implements HttpHandler {
        public void handle(HttpExchange ex) throws IOException {
            if (handleOptions(ex)) return;
            try {
                Connection conn = getDb();
                synchronized (DB_LOCK) {
                    int wc=0, rc=0;
                    ResultSet r1 = conn.createStatement().executeQuery("SELECT COUNT(*) FROM words");
                    if (r1.next()) wc = r1.getInt(1); r1.close();
                    ResultSet r2 = conn.createStatement().executeQuery("SELECT COUNT(*) FROM roots");
                    if (r2.next()) rc = r2.getInt(1); r2.close();
                    sendJson(ex, 200, "{\"hasData\":" + (wc > 0 || rc > 0) + "}");
                }
            } catch (Exception e) { sendJson(ex, 500, "{\"error\":\"" + esc(e.getMessage()) + "\"}"); }
        }
    }

    // ===== InitWordsHandler =====
    static class InitWordsHandler implements HttpHandler {
        public void handle(HttpExchange ex) throws IOException {
            if (handleOptions(ex)) return;
            try {
                String body = readBody(ex);
                List<String> items; body = body.trim();
                if (body.startsWith("[")) { items = jsonArray(body); }
                else { String arr = jsonStr(body,"words"); items = arr.isEmpty() ? new ArrayList<String>() : jsonArray(arr); }
                Connection conn = getDb();
                synchronized (DB_LOCK) {
                    PreparedStatement ps = conn.prepareStatement("INSERT INTO words(cn,en,cat,roots,score,abbr,cnDesc,enDesc,ref,status,time) VALUES(?,?,?,?,?,?,?,?,?,?,?)");
                    for (String item : items) {
                        ps.setString(1,jsonStr(item,"cn")); ps.setString(2,jsonStr(item,"en")); ps.setString(3,jsonStr(item,"cat"));
                        ps.setString(4,jsonStr(item,"roots")); ps.setDouble(5,jsonNum(item,"score")); ps.setString(6,jsonStr(item,"abbr"));
                        ps.setString(7,jsonStr(item,"cnDesc")); ps.setString(8,jsonStr(item,"enDesc")); ps.setString(9,jsonStr(item,"ref"));
                        String st = jsonStr(item,"status"); ps.setString(10, st.isEmpty()?"approved":st); ps.setString(11,jsonStr(item,"time"));
                        ps.addBatch();
                    }
                    ps.executeBatch(); ps.close();
                }
                writeLog("\u2705 \u521D\u59CB\u5316\u8BCD\u6761 - \u5171" + items.size() + "\u6761");
                writeLog(reqLine(ex));
                sendJson(ex, 200, "{\"msg\":\"ok\",\"count\":" + items.size() + "}");
            } catch (Exception e) { writeLog("\u274C \u521D\u59CB\u5316\u8BCD\u6761\u5F02\u5E38: " + e.getMessage()); sendJson(ex, 500, "{\"error\":\"" + esc(e.getMessage()) + "\"}"); }
        }
    }

    // ===== InitRootsHandler =====
    static class InitRootsHandler implements HttpHandler {
        public void handle(HttpExchange ex) throws IOException {
            if (handleOptions(ex)) return;
            try {
                String body = readBody(ex); body = body.trim();
                List<String> items;
                if (body.startsWith("[")) { items = jsonArray(body); }
                else { String arr = jsonStr(body,"roots"); items = arr.isEmpty() ? new ArrayList<String>() : jsonArray(arr); }
                Connection conn = getDb();
                synchronized (DB_LOCK) {
                    PreparedStatement ps = conn.prepareStatement("INSERT INTO roots(name,en,mean,src,cat,status,examples) VALUES(?,?,?,?,?,?,?)");
                    for (String item : items) {
                        ps.setString(1,jsonStr(item,"name")); ps.setString(2,jsonStr(item,"en")); ps.setString(3,jsonStr(item,"mean"));
                        ps.setString(4,jsonStr(item,"src")); ps.setString(5,jsonStr(item,"cat"));
                        String st = jsonStr(item,"status"); ps.setString(6, st.isEmpty()?"approved":st); ps.setString(7,jsonStr(item,"examples"));
                        ps.addBatch();
                    }
                    ps.executeBatch(); ps.close();
                }
                writeLog("\u2705 \u521D\u59CB\u5316\u8BCD\u6839 - \u5171" + items.size() + "\u6761");
                writeLog(reqLine(ex));
                sendJson(ex, 200, "{\"msg\":\"ok\",\"count\":" + items.size() + "}");
            } catch (Exception e) { writeLog("\u274C \u521D\u59CB\u5316\u8BCD\u6839\u5F02\u5E38: " + e.getMessage()); sendJson(ex, 500, "{\"error\":\"" + esc(e.getMessage()) + "\"}"); }
        }
    }

    // ===== ImportCheckHandler =====
    static class ImportCheckHandler implements HttpHandler {
        public void handle(HttpExchange ex) throws IOException {
            if (handleOptions(ex)) return;
            try {
                String body = readBody(ex);
                List<String> items = jsonArray(body);
                int total = items.size();
                // 自比对：检查导入数据内部cn+en重复
                Set<String> seen = new HashSet<>();
                List<String> selfDups = new ArrayList<>();
                for (String item : items) {
                    String key = jsonStr(item,"cn").trim() + "|" + jsonStr(item,"en").trim().toLowerCase();
                    if (seen.contains(key)) {
                        selfDups.add("{\"cn\":\"" + esc(jsonStr(item,"cn")) + "\",\"en\":\"" + esc(jsonStr(item,"en")) + "\"}");
                    } else { seen.add(key); }
                }
                // 平台比对：加载所有现有词�?
                Connection conn = getDb();
                List<String[]> existing = new ArrayList<>();
                synchronized (DB_LOCK) {
                    ResultSet rs = conn.createStatement().executeQuery("SELECT cn,en FROM words");
                    while (rs.next()) existing.add(new String[]{rs.getString("cn"), rs.getString("en")});
                    rs.close();
                }
                List<String> exactDups = new ArrayList<>();
                List<String> similarItems = new ArrayList<>();
                for (String item : items) {
                    String cn = jsonStr(item,"cn").trim(), en = jsonStr(item,"en").trim();
                    for (String[] e : existing) {
                        if (cn.equals(e[0]) && en.equalsIgnoreCase(e[1])) {
                            exactDups.add("{\"cn\":\"" + esc(cn) + "\",\"en\":\"" + esc(en) + "\"}");
                            break;
                        }
                    }
                    for (String[] e : existing) {
                        if (cn.equals(e[0]) && en.equalsIgnoreCase(e[1])) continue;
                        double simCn = calcSimilarity(cn, e[0]);
                        double simEn = calcSimilarity(en, e[1]);
                        double sim = Math.max(simCn, simEn);
                        if (sim >= 0.7) {
                            similarItems.add("{\"importCn\":\"" + esc(cn) + "\",\"importEn\":\"" + esc(en) +
                                "\",\"existCn\":\"" + esc(e[0]) + "\",\"existEn\":\"" + esc(e[1]) +
                                "\",\"similarity\":" + String.format("%.2f", sim) + "}");
                            break;
                        }
                    }
                }
                String result = "{\"total\":" + total +
                    ",\"selfDupCount\":" + selfDups.size() +
                    ",\"exactDupCount\":" + exactDups.size() +
                    ",\"similarCount\":" + similarItems.size() +
                    ",\"selfDups\":[" + String.join(",", selfDups) + "]" +
                    ",\"exactDups\":[" + String.join(",", exactDups) + "]" +
                    ",\"similarItems\":[" + String.join(",", similarItems) + "]}";
                sendJson(ex, 200, result);
                writeLog("\u2705 \u5BFC\u5165\u6BD4\u5BF9 - \u5171" + total + "\u6761, \u5B8C\u5168\u91CD\u590D:" + exactDups.size() + "\u6761, \u76F8\u4F3C:" + similarItems.size() + "\u6761");
                writeLog(reqLine(ex));
            } catch (Exception e) {
                writeLog("\u274C \u5BFC\u5165\u6BD4\u5BF9\u5F02\u5E38: " + e.getMessage());
                sendJson(ex, 500, "{\"error\":\"" + esc(e.getMessage()) + "\"}");
            }
        }
    }

    // ===== StaticHandler =====
    static class StaticHandler implements HttpHandler {
        public void handle(HttpExchange ex) throws IOException {
            String path = ex.getRequestURI().getPath();
            if ("/".equals(path) || path.isEmpty()) path = "/index.html";
            File f = new File(SCRIPT_DIR, path.substring(1));
            if (!f.exists() || f.isDirectory()) {
                f = new File(SCRIPT_DIR, "index.html");
            }
            if (!f.exists()) {
                String msg = "Not Found";
                ex.sendResponseHeaders(404, msg.length());
                ex.getResponseBody().write(msg.getBytes());
                ex.getResponseBody().close();
                return;
            }
            String ct = "text/html; charset=utf-8";
            if (path.endsWith(".js")) ct = "application/javascript; charset=utf-8";
            else if (path.endsWith(".css")) ct = "text/css; charset=utf-8";
            else if (path.endsWith(".json")) ct = "application/json; charset=utf-8";
            else if (path.endsWith(".png")) ct = "image/png";
            else if (path.endsWith(".jpg") || path.endsWith(".jpeg")) ct = "image/jpeg";
            else if (path.endsWith(".svg")) ct = "image/svg+xml";
            else if (path.endsWith(".ico")) ct = "image/x-icon";
            ex.getResponseHeaders().set("Content-Type", ct);
            addCors(ex);
            byte[] data = Files.readAllBytes(f.toPath());
            ex.sendResponseHeaders(200, data.length);
            ex.getResponseBody().write(data);
            ex.getResponseBody().close();
        }
    }

}
