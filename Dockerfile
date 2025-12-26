FROM ubuntu:latest AS build
RUN apt-get update
RUN apt-get install openjdk-11-jdk -y
COPY . .
FROM openjdk:11-jdk-slim

EXPOSE 8080
ARG JAR_FILE=./target/*.jar
COPY ${JAR_FILE} app.jar
ENTRYPOINT ["java", "-jar", "app.jar"]
