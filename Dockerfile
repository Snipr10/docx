FROM gradle:7.1.0-jdk11 AS TEMP_BUILD_IMAGE
ENV APP_HOME=.
WORKDIR $APP_HOME
COPY build.gradle $APP_HOME

COPY gradle $APP_HOME/gradle
USER root
RUN gradle compileJava || return 0
RUN gradle build || return 0
COPY . .
ENTRYPOINT gradle run


