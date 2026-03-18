FROM registry.yygu.cn/library/dotnet_sdk:8.0 AS build
WORKDIR /src

COPY ["Ofd2PdfService.csproj", "./"]
RUN dotnet restore "Ofd2PdfService.csproj"

COPY . .
RUN dotnet publish "Ofd2PdfService.csproj" -c Release -o /app/publish --no-restore

FROM registry.yygu.cn/library/dotnet_aspnet:8.0 AS runtime
WORKDIR /app

# Install native font/image dependencies and Chinese fonts required by Spire.PDF on Linux
RUN apt-get update && apt-get install -y --no-install-recommends \
    libfontconfig1 \
    libgdiplus \
    libc6-dev \
    fonts-noto-cjk \
    fonts-wqy-zenhei \
    fonts-wqy-microhei \
    && fc-cache -fv \
    && rm -rf /var/lib/apt/lists/*

COPY --from=build /app/publish .

ENV ASPNETCORE_URLS=http://+:8080
EXPOSE 8080

ENTRYPOINT ["dotnet", "Ofd2PdfService.dll"]
