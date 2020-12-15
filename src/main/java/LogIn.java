import io.undertow.server.HttpHandler;
import io.undertow.server.HttpServerExchange;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.net.*;
import java.util.List;

public class LogIn implements HttpHandler {
    CookieManager cookieManager;

    @Override
    public void handleRequest(HttpServerExchange exchange) throws Exception {

        cookieManager = new CookieManager(null, CookiePolicy.ACCEPT_ALL);
        CookieHandler.setDefault(cookieManager);
        String jsonInputString = "{\"login\": \"java_api\", \"password\": \"4yEcwVnjEH7D\"}";
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
        for(Object o: (JSONArray)(jsonPosts).get("total")){
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
        WordWorker.createDoc("«Гусев Олег Александрович»" , "1 июля - 31 июля 2020 года", data, jsonPosts, jsonComments);

        exchange.getResponseSender().send("OK");
    }

    private String get_info() throws IOException {
        URL url = new URL("https://api.glassen-it.com/component/socparser/threads/get");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);

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
        return res;
    }

    private JSONObject getCommentInfo()throws IOException {
        URL url = new URL("https://api.glassen-it.com/component/socparser/stats/commentTrustDaily");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        String jsonInputString = "{\"thread_id\": \"995\", \"from\": \"2010-01-01\", \"to\": \"2020-12-20\"}";
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
        String jsonInputString = "{\"thread_id\": \"995\", \"from\": \"2010-01-01\", \"to\": \"2020-12-20\"}";
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
//        JSONArray jsonArray;
        return (JSONObject)jsonObjects.get("total");
//        for(Object o: (JSONArray)((JSONObject)jsonObjects.get("total")).get("total")){
//            jsonArray = (JSONArray) o;
//            total += (int) jsonArray.get(1);
//        }
//        return total;
    }
    private Integer getPostsSources() throws IOException {
        URL url = new URL("https://api.glassen-it.com/component/socparser/stats/thread/allmembers");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        String jsonInputString = "{\"thread_id\": \"995\", \"from\": \"2010-01-01\", \"to\": \"2020-12-20\", \"sources_only\": \"true\"}";
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
        int total = 0;
        JSONObject jsonObjects = new JSONObject(res);
        JSONArray jsonArray;
        for(Object o: (JSONArray)jsonObjects.get("graph_data")){
            jsonArray = (JSONArray) o;
            total += (int) jsonArray.get(1);
        }
        return total;
    }

}
