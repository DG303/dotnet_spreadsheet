# Build stage
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Copy the project files
COPY ["src/SpreadsheetEditor.Console/SpreadsheetEditor.Console.csproj", "src/SpreadsheetEditor.Console/"]
COPY ["src/SpreadsheetEditor.Core/SpreadsheetEditor.Core.csproj", "src/SpreadsheetEditor.Core/"]

# Restore dependencies
RUN dotnet restore "src/SpreadsheetEditor.Console/SpreadsheetEditor.Console.csproj"

# Copy the rest of the code
COPY . .

# Build the application
RUN dotnet build "src/SpreadsheetEditor.Console/SpreadsheetEditor.Console.csproj" -c Release -o /app/build

# Publish the application
RUN dotnet publish "src/SpreadsheetEditor.Console/SpreadsheetEditor.Console.csproj" -c Release -o /app/publish

# Final stage
FROM mcr.microsoft.com/dotnet/runtime:8.0
WORKDIR /app

# Copy the published application
COPY --from=build /app/publish .

# Create a directory for Excel files
RUN mkdir -p /app/data

# Set the entry point
ENTRYPOINT ["dotnet", "SpreadsheetEditor.Console.dll"] 