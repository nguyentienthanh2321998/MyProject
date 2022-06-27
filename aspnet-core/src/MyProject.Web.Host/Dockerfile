#FROM mcr.microsoft.com/dotnet/core/aspnet:2.2-stretch-slim

#WORKDIR /app
#COPY . .

#ENTRYPOINT ["dotnet", "MyProject.Web.Host.dll"]


# Dockerfile

FROM mcr.microsoft.com/dotnet/core/aspnet:2.2-stretch-slim AS build-env
WORKDIR /app

# Copy csproj and restore as distinct layers
COPY *.csproj ./
RUN dotnet restore

# Copy everything else and build
COPY . .
RUN dotnet publish -c Release -o out

# Build runtime image
FROM mcr.microsoft.com/dotnet/core/aspnet:2.2-stretch-slim
WORKDIR /app
COPY --from=build-env /app/out .

# Run the app on container startup
# Use your project name for the second parameter
# e.g. MyProject.dll
ENTRYPOINT ["dotnet", "MyProject.Web.Host.dll"]