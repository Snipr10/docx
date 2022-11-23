ARG BUILD_HOME=/report
FROM gradle:7.1.0-jdk11  as build-image
ARG BUILD_HOME
ENV APP_HOME=$BUILD_HOME
WORKDIR $APP_HOME
COPY --chown=gradle:gradle .env build.gradle settings.gradle $APP_HOME/
COPY --chown=gradle:gradle src $APP_HOME/src

# Build the application.
RUN gradle build || return 0
FROM adoptopenjdk/openjdk16
RUN apt-get install libfreetype6
ARG BUILD_HOME
ENV APP_HOME=$BUILD_HOME
COPY --from=build-image $APP_HOME/build/libs/report.jar app.jar
COPY --from=build-image $APP_HOME/.env .env

ENTRYPOINT java -jar app.jar
