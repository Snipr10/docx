import io.undertow.Handlers;
import io.undertow.Undertow;

import java.io.IOException;

//@lombok.extern.slf4j.Slf4j
public class Main {
    public static void main(String[] args) throws IOException {

        final
        Undertow server = Undertow.builder()
                .addHttpListener(4274, "0.0.0.0")

                .setHandler(
                        Handlers.path()
                                .addExactPath("/data",
                                        Handlers.routing()
                                                .get("/", new Data()))
                                .addExactPath("/test",
                                        Handlers.routing()
                                                .get("/", new Test()))
                                .addExactPath("/login",
                                        Handlers.routing()
                                                .get("/", new LogIn()))
                )

                .build();
//        log.info("Server start port 4274");
        server.start();
        System.out.println("start");
    }
}
