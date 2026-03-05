# Base stage for running the app
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
USER $APP_UID
WORKDIR /app
EXPOSE 8080
EXPOSE 8081

# Build stage
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
ARG BUILD_CONFIGURATION=Release
WORKDIR /src

# 1. COPY using the full paths from the REPO ROOT
# Note the 'src/' prefix added to these paths
COPY ["src/doclayer_webapi/doclayer_webapi/doclayer_webapi.csproj", "src/doclayer_webapi/doclayer_webapi/"]
COPY ["src/DocLayer.Core/DocLayer.Core/DocLayer.Core.csproj", "src/DocLayer.Core/DocLayer.Core/"]

# 2. Restore
RUN dotnet restore "src/doclayer_webapi/doclayer_webapi/doclayer_webapi.csproj"

# 3. Copy everything (Respects .dockerignore)
COPY . .

# 4. Build
WORKDIR "/src/src/doclayer_webapi/doclayer_webapi"
RUN dotnet build "./doclayer_webapi.csproj" -c $BUILD_CONFIGURATION -o /app/build

# Publish stage
FROM build AS publish
ARG BUILD_CONFIGURATION=Release
RUN dotnet publish "./doclayer_webapi.csproj" -c $BUILD_CONFIGURATION -o /app/publish /p:UseAppHost=false

# Final stage
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "doclayer_webapi.dll"]