import io.undertow.server.HttpHandler;
import io.undertow.server.HttpServerExchange;
import io.undertow.util.Headers;
import io.undertow.util.HttpString;
import org.apache.poi.xwpf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.net.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class LogIn implements HttpHandler {
    CookieManager cookieManager;
    String dateFrom;
    String dateTo;
    String thread_id;
    String type;

    public void handleRequest(HttpServerExchange exchange) throws Exception {
        if (exchange.isInIoThread()) {
            exchange.dispatch(this);
            return;
        }
        exchange.startBlocking();

        BufferedReader reader = null;
        reader = new BufferedReader( new InputStreamReader( exchange.getInputStream( ) ) );
        StringBuilder json = new StringBuilder();
        String line;
        while( ( line = reader.readLine( ) ) != null ) {
            json.append(line);
        }
        JSONObject jsonObject = new JSONObject(json.toString());

        Map<String, Deque<String>> params = exchange.getQueryParameters();
        dateFrom = params.get("dateFrom").getFirst();
        dateTo = params.get("dateTo").getFirst();
        thread_id = params.get("thread_id").getFirst();
        type = params.get("type").getFirst();
        Date dateFromReal = new SimpleDateFormat("yyyy-MM-dd").parse(dateFrom);
        Date dateToReal = new SimpleDateFormat("yyyy-MM-dd").parse(dateTo);

        Calendar cal = Calendar.getInstance(TimeZone.getTimeZone("Europe/Paris"));
        cal.setTime(dateFromReal);
        int first_month = cal.get(Calendar.MONTH);
        int first_year = cal.get(Calendar.YEAR);

        String dateFromString =
                DateFormat.getDateInstance(SimpleDateFormat.LONG, new Locale("ru")).format(dateFromReal)
                .replace(cal.get(Calendar.YEAR) + " г.", "");
        cal.setTime(dateFromReal);
        String yearFrom = String.valueOf(cal.get(Calendar.YEAR));

        cal.setTime(dateToReal);
        String year = String.valueOf(cal.get(Calendar.YEAR));
        String dateToString =
                DateFormat.getDateInstance(SimpleDateFormat.LONG, new Locale("ru")).format(dateToReal)
                        .replace(year + " г.", "");
        dateFromReal. toInstant() .toString();

        dateFrom += " 00:00:00";
        dateTo += " 23:59:59";
        cookieManager = new CookieManager(null, CookiePolicy.ACCEPT_ALL);
        CookieHandler.setDefault(cookieManager);
        String jsonInputString = String.format("{\"login\": \"%s\", \"password\": \"%s\"}",
                jsonObject.get("login"), jsonObject.get("password"));
        URL url = new URL("https://api.glassen-it.com/component/socparser/authorization/login");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("GET");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            System.out.println(response.toString());
        }
        List<HttpCookie> cookies = cookieManager.getCookieStore().getCookies();
        for (HttpCookie cookie : cookies) {
            System.out.println(cookie.getDomain());
            System.out.println(cookie);
        }
        url = new URL("https://api.glassen-it.com/component/socparser/authorization/login");
        connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }

        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            System.out.println(response.toString());
        }

        int total_sources = getPostsSources();
        JSONObject jsonPosts = getPostsInfo();
        JSONArray jsonArray;
        int total_publication = 0;

        for(Object o: (JSONArray)((JSONObject)jsonPosts.get("total")).get("total")){
            jsonArray = (JSONArray) o;
            total_publication += (int) jsonArray.get(1);
        }
        int total_comment = 0;
        JSONObject jsonComments = getCommentInfo();
        for(Object o: (JSONArray)(jsonComments).get("total")){
            jsonArray = (JSONArray) o;
            total_comment += (int) jsonArray.get(1);
        }
        DataForDocx data = new DataForDocx(total_sources, total_publication, total_comment);
        JSONArray postsContent =getPostsContent();
        JSONArray commentContent = getCommentContent();
        JSONArray posts = getPosts();
        JSONObject stat = getStats();
        JSONObject sex = getStats("sex");
        JSONObject age = getAge();
        JSONArray jsonCity = getCity();
        JSONObject usersJson = getUsers();

        XWPFDocument docx = WordWorker.createDoc(type, getNameThread(), String.format("%s%s года - %s %s года", dateFromString, yearFrom, dateToString, year),
                data, jsonPosts, jsonComments, stat, sex, age, usersJson, jsonCity, posts, postsContent, commentContent,
                first_month, first_year
                );
        final String name = UUID.randomUUID() + ".docx";
        try (FileOutputStream fileOut = new FileOutputStream(name)) {
            docx.write(fileOut);
            }

        exchange.getResponseHeaders().put(Headers.CONTENT_TYPE,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        exchange.getResponseHeaders().put(Headers.CONTENT_DISPOSITION, "attachment; filename=\"" + name+"\"");
        exchange.getResponseHeaders().put(new HttpString("Access-Control-Allow-Methods"),
                "GET, POST, PUT, DELETE, OPTIONS");
        exchange.getResponseHeaders()
                .put(new HttpString("Access-Control-Allow-Origin"), "*");
        exchange.getResponseHeaders()
                .put(new HttpString("Content-Description"), "File Transfer");

        exchange.getResponseHeaders()
                .put(new HttpString("Content-Transfer-Encoding"), "binary");




        exchange.getResponseHeaders()
                .put(new HttpString("Pragma"), "public");
        final File file = new File(name);
        final OutputStream outputStream = exchange.getOutputStream();
        final InputStream inputStream = new FileInputStream(file);
        int length = inputStream.available();
        exchange.getResponseHeaders()
                .put(new HttpString("Content-Length"), length);
        byte[] buf = new byte[8192];
        int c;
        while ((c = inputStream.read(buf, 0, buf.length)) > 0) {
            outputStream.write(buf, 0, c);
            outputStream.flush();
        }


        outputStream.close();
        inputStream.close();
        exchange.getResponseSender().send("OK");
    }


    private JSONObject getCommentInfo()throws IOException {
        URL url = new URL("https://api.glassen-it.com/component/socparser/stats/commentTrustDaily");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        String jsonInputString = String.format("{\"thread_id\": \"%s\", \"from\": \"%s\", \"to\": \"%s\"}",
                thread_id, dateFrom, dateTo);
        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }
        String res = "";
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            res += response.toString();
        }
        return (JSONObject)(new JSONObject(res)).get("total");
    }

    private JSONObject getPostsInfo()throws IOException {
        URL url = new URL("https://api.glassen-it.com/component/socparser/stats/trustdaily");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        String jsonInputString = String.format("{\"thread_id\": \"%s\", \"from\": \"%s\", \"to\": \"%s\"}",
                thread_id, dateFrom, dateTo);
        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }
        String res = "";
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            res += response.toString();
        }
        JSONObject jsonObjects = new JSONObject(res);
        return jsonObjects;

    }
    private Integer getPostsSources() throws IOException {
        URL url = new URL("https://api.glassen-it.com/component/socparser/content/membersCount");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        String jsonInputString = String.format("{\"thread_id\": \"%s\", \"from\": \"%s\", \"to\": \"%s\"}",
                thread_id, dateFrom, dateTo);
        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }
        String res = "";
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            res += response.toString();
        }
        return (Integer) new JSONObject(res).get("source_count");
    }



//    private JSONObject getUsers() throws IOException {
//        URL url = new URL("https://api.glassen-it.com/component/socparser/thread/allmembers");
//        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
//        connection.setRequestMethod("POST");
//        connection.setRequestProperty("Content-Type", "application/json; utf-8");
//        connection.setRequestProperty("Accept", "application/json");
//        connection.setDoOutput(true);
//        String jsonInputString = String.format(
//                "{\"thread_id\": \"%s\", \"from\": \"%s\", \"to\": \"%s\", \"sources_only\": \"0\", \"limit\": \"10\"}",
//                        thread_id, dateFrom, dateTo);
//        try (OutputStream os = connection.getOutputStream()) {
//            byte[] input = jsonInputString.getBytes("utf-8");
//            os.write(input, 0, input.length);
//            os.flush();
//        }
//        String res = "";
//        try (BufferedReader br = new BufferedReader(
//                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
//            StringBuilder response = new StringBuilder();
//            String responseLine = null;
//            while ((responseLine = br.readLine()) != null) {
//                response.append(responseLine.trim());
//            }
//            res += response.toString();
//        }
//         return new JSONObject(res);
//
//    }

        private JSONObject getUsers() throws IOException {
        URL url = new URL("https://api.glassen-it.com/component/socparser/stats/userlinks ");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        String jsonInputString = String.format(
                "{\"thread_id\": \"%s\", \"from\": \"%s\", \"to\": \"%s\", \"start\": \"0\", \"limit\": \"10\"}",
                        thread_id, dateFrom, dateTo);
        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }
        String res = "";
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            res += response.toString();
        }
         return new JSONObject(res);

    }

    private JSONObject getStats()throws IOException {
        URL url = new URL("https://api.glassen-it.com/component/socparser/stats");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        String jsonInputString = String.format("{\"thread_id\": \"%s\", \"from\": \"%s\", \"to\": \"%s\"}",
                thread_id, dateFrom, dateTo);
        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }
        String res = "";
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            res += response.toString();
        }
        JSONObject jsonObjects = new JSONObject(res);
        return jsonObjects;

    }

    private JSONArray getCity()throws IOException {
        URL url = new URL("https://api.glassen-it.com/component/socparser/thread/getcitytop");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        String jsonInputString = String.format("{\"thread_id\": \"%s\", \"limit\": \"10\"}",
                thread_id, dateFrom, dateTo);
        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }
        String res = "";
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            res += response.toString();
        }
        return new JSONArray(res);

    }

    private JSONObject getStats(String type)throws IOException {
        URL url = new URL("https://api.glassen-it.com/component/socparser/stats");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        String jsonInputString = String.format(
                "{\"thread_id\": \"%s\", \"from\": \"%s\", \"to\": \"%s\", \"type\":\"%s\"}",
                thread_id, dateFrom, dateTo,type);

        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }
        String res = "";
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            res += response.toString();
        }
        JSONObject jsonObjects = new JSONObject(res);
        return jsonObjects;

    }

    private JSONObject getAge()throws IOException {
        URL url = new URL("https://api.glassen-it.com/component/socparser/stats/ages");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        String jsonInputString = String.format(

        "{\"thread_id\": \"%s\", \"from\": \"%s\", \"to\": \"%s\", \"group1_start\":\"18\",\"group1_end\":\"25\",\"group2_start\":\"25\",\"group2_end\":\"40\",\"group3_start\":\"40\",\"group3_end\":\"200\",\"group4_start\":\"0\",\"group4_end\":\"0\"}",
                thread_id, dateFrom, dateTo);

        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }
        String res = "";
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            res += response.toString();
        }

        return (JSONObject)((JSONObject)new JSONObject(res).get("additional_data")).get("age");

    }

    private JSONArray getPosts()throws IOException {
        URL url = new URL("https://api.glassen-it.com/component/socparser/stats/owners_top");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        String jsonInputString = String.format("{\"thread_id\": \"%s\", \"from\": \"%s\", \"to\": \"%s\", \"limit\": \"10\"}",
                thread_id, dateFrom, dateTo);
        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }
        String res = "";
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            res += response.toString();
        }
        JSONArray jsonObjects = new JSONArray(res);
        return jsonObjects;

    }


    private JSONArray getPostsContent() throws IOException {
        URL url = new URL("https://api.glassen-it.com/component/socparser/content/posts");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        String jsonInputString =String.format(
                "{\"thread_id\": \"%s\", \"from\": \"%s\", \"to\": \"%s\", \"limit\":\"10\" ,  " +
                        "\"sort\": {\"type\": \"viewed\",\"order\": \"desc\"}}",
                thread_id, dateFrom, dateTo);

        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }
        String res = "";
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            res += response.toString();
        }
        return (JSONArray) new JSONObject(res).get("posts");

    }

    private JSONArray getCommentContent() throws IOException {
        URL url = new URL("https://api.glassen-it.com/component/socparser/content/allcommentaries");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        String jsonInputString =String.format(
                "{\"thread_id\": \"%s\", \"from\": \"%s\", \"to\": \"%s\", \"limit\":\"10\" ,  " +
                        "\"sort\": {\"type\": \"likes\",\"order\": \"desc\"}}",
        thread_id, dateFrom, dateTo);

        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }
        String res = "";
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            res += response.toString();
        }
        return (JSONArray) new JSONObject(res).get("commentaries");

    }

    private String getNameThread() throws IOException{
        URL url = new URL("https://api.glassen-it.com/component/socparser/thread/additional_info");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        String jsonInputString =String.format(
                "{\"thread_id\": \"%s\"}",
                thread_id);
        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }
        String res = "";
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            res += response.toString();
        }
        return (String) new JSONObject(res).get("name");

    }

}
