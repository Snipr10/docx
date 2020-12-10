import io.undertow.server.HttpHandler;
import io.undertow.server.HttpServerExchange;

public class Test implements HttpHandler {
    @Override
    public void handleRequest(HttpServerExchange exchange) {
        WordWorker.createDoc("«Гусев Олег Александрович»" , "1 июля - 31 июля 2020 года");
        exchange.getResponseSender().send("Read me");
    }

}
