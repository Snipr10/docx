import io.undertow.Handlers;
import io.undertow.Undertow;
import io.undertow.server.HttpHandler;
import io.undertow.server.HttpServerExchange;
import io.undertow.util.HttpString;
import io.github.cdimascio.dotenv.Dotenv;

import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {

        final
        Undertow server = Undertow.builder()
                .addHttpListener(4274, "0.0.0.0")

//                .setHandler(
//                        Handlers.path()
//                                .addExactPath("/data",
//                                        Handlers.routing()
//                                                .post("/", new LogIn())),
//                )
                .setHandler(new HttpHandler() {
                    @Override
                    public void handleRequest(HttpServerExchange exchange) throws Exception {
                        Dotenv dotenv = null;
                        dotenv = Dotenv.configure().load();
                        if (exchange.getRequestMethod().toString().equals("POST")) {
                            LogIn s = new LogIn(dotenv.get("DOMAIN"));

                            s.handleRequest(exchange);
                        }
                        else {
                            exchange.getResponseHeaders().put(new HttpString("Access-Control-Allow-Methods"),
                                    "*");
                            exchange.getResponseHeaders()
                                    .put(new HttpString("Access-Control-Allow-Headers"),
                                            "origin, x-requested-with, accept, accept-language, content-language, content-type, Access-Control-Request-Headers, Access-Control-Request-Method");
                            exchange.getResponseHeaders().put(new HttpString("Access-Control-Max-Age"),
                                    "1728000");
                            exchange.getResponseHeaders().put(new HttpString("Content-Length"),
                                    "0");
                            exchange.getResponseHeaders().put(new HttpString("Content-Type"),
                                    "text/plain");

                        }
                    }
                })
                .build();
        server.start();
        System.out.println("start");
    }
}
