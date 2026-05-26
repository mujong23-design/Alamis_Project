# ================================================
# Multi-stage build: Maven 빌드 → 실행 전용 경량 이미지
# ================================================

# ---- Stage 1: Build ----
FROM maven:3.9-eclipse-temurin-17 AS build
WORKDIR /app

# 1) pom.xml 만 먼저 복사하여 의존성 캐싱
COPY pom.xml .
RUN mvn -B -q dependency:go-offline

# 2) 소스 복사 후 빌드 (테스트는 배포 빌드에서 스킵)
COPY src ./src
RUN mvn -B -q clean package -DskipTests

# ---- Stage 2: Runtime ----
FROM eclipse-temurin:17-jre
WORKDIR /app

# 빌드 산출물만 복사 (WAR 패키징 - JSP 정상 동작)
COPY --from=build /app/target/inventory-*.war app.war

# Render 등 클라우드에서 PORT 환경변수 주입 → Spring 이 ${PORT:8080} 으로 받음
EXPOSE 8080

# Render 512MB RAM 대비 힙 캡(380MB) + UTF-8 인코딩 고정
ENTRYPOINT ["java","-Xmx380m","-Dfile.encoding=UTF-8","-jar","app.war"]
