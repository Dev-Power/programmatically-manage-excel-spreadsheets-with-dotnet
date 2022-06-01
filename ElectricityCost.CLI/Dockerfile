FROM mcr.microsoft.com/dotnet/runtime:6.0 AS base
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /src
COPY ["ElectricityCost.CLI/ElectricityCost.CLI.csproj", "ElectricityCost.CLI/"]
RUN dotnet restore "ElectricityCost.CLI/ElectricityCost.CLI.csproj"
COPY . .
WORKDIR "/src/ElectricityCost.CLI"
RUN dotnet build "ElectricityCost.CLI.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "ElectricityCost.CLI.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "ElectricityCost.CLI.dll"]
