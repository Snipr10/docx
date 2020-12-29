import io.undertow.Handlers;
import io.undertow.Undertow;

import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {

        final
        Undertow server = Undertow.builder()
                .addHttpListener(4274, "0.0.0.0")

                .setHandler(
                        Handlers.path()
                                .addExactPath("/data",
                                        Handlers.routing()
                                                .post("/", new LogIn()))
                )

                .build();
        server.start();
        System.out.println("start");
    }
}
