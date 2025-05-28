FROM mcr.microsoft.com/dotnet/sdk:9.0 AS build
WORKDIR /src

# Copy everything from the current directory (which contains .csproj, code, etc.)
COPY . .

# Restore dependencies and publish the app to /app/publish
RUN dotnet restore
RUN dotnet publish -c Release -o /app/publish /p:PublishReadyToRun=false

# Final image: Lambda runtime base for .NET 9
FROM public.ecr.aws/lambda/dotnet:9
WORKDIR /var/task

# Copy published output from build stage
COPY --from=build /app/publish .

# Set the Lambda handler
CMD [ "WordDocProcessor::WordDocProcessor.Function::FunctionHandler" ]
