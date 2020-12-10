import io.undertow.server.HttpHandler;
import io.undertow.server.HttpServerExchange;
import org.apache.commons.lang3.StringUtils;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.*;
import java.util.List;
import java.util.Map;

public class LogIn implements HttpHandler {
    @Override
    public void handleRequest(HttpServerExchange exchange) throws IOException {
        CookieManager cookieManager = new CookieManager(null, CookiePolicy.ACCEPT_ALL);
        CookieHandler.setDefault(cookieManager);
        String jsonInputString = "{\"login\": \"java_api\", \"password\": \"4yEcwVnjEH7D\"}";
        URL url = new URL("https://api.glassen-it.com/component/socparser/authorization/login");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("GET");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        try(OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }
        try(BufferedReader br = new BufferedReader(
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
        try(OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
            os.flush();
        }

        try(BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            System.out.println(response.toString());
        }
        url = new URL("https://api.glassen-it.com/component/socparser/threads/get");
        connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json; utf-8");
        connection.setRequestProperty("Accept", "application/json");
        connection.setDoOutput(true);
        if (cookies != null) {
            if (cookies.size() > 0) {
//                Debug("{0} -- Adding Cookie Headers : ", url.toString());
//                for (HttpCookie cookie : cookies) {
//                    Debug(cookie.toString(), null);
//                }

                //adding the cookie header
                connection.setRequestProperty("Cookie", StringUtils.join(cookies, ";"));
            }
        }


        try(OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonInputString.getBytes("utf-8");
            os.write(input, 0, input.length);
        }
        String res = "";
        try(BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine = null;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            res += response.toString();
//            System.out.println(response.toString());
        }
//        res = "";
//
//        url = new URL("https://api.glassen-it.com/component/socparser/users/get");
//        connection = (HttpURLConnection) url.openConnection();
//        connection.setRequestMethod("POST");
//        connection.setRequestProperty("Content-Type", "application/json; utf-8");
//        connection.setRequestProperty("Accept", "application/json");
//        connection.setDoOutput(true);
//        if (cookies != null) {
//            if (cookies.size() > 0) {
////                Debug("{0} -- Adding Cookie Headers : ", url.toString());
////                for (HttpCookie cookie : cookies) {
////                    Debug(cookie.toString(), null);
////                }
//
//                //adding the cookie header
//                connection.setRequestProperty("Cookie", StringUtils.join(cookies, ";"));
//            }
//        }
//
//
//        try(OutputStream os = connection.getOutputStream()) {
//            byte[] input = jsonInputString.getBytes("utf-8");
//            os.write(input, 0, input.length);
//        }
//        try(BufferedReader br = new BufferedReader(
//                new InputStreamReader(connection.getInputStream(), "utf-8"))) {
//            StringBuilder response = new StringBuilder();
//            String responseLine = null;
//            while ((responseLine = br.readLine()) != null) {
//                response.append(responseLine.trim());
//            }
//            res += response.toString();
////            System.out.println(response.toString());
//        }
        exchange.getResponseSender().send(res);
    }

}
