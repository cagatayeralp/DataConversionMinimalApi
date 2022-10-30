#See https://aka.ms/containerfastmode to understand how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/aspnet:6.0 AS base
WORKDIR /app
EXPOSE 8001
EXPOSE 443

ENV ASPNETCORE_URLS="http://*:5000"
ENV ASPNETCORE_ENVIRONMENT="Development"

FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /src
COPY ["DataConversion.csproj", "."]
RUN dotnet restore "./DataConversion.csproj"
COPY . .
WORKDIR "/src/."
RUN dotnet build "DataConversion.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "DataConversion.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "DataConversion.dll"]