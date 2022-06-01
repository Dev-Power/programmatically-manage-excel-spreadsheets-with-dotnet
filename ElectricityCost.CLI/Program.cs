using CliFx;

public static class Program
{
    public static async Task<int> Main() =>
        await new CliApplicationBuilder()
            .AddCommandsFromThisAssembly()
            .SetExecutableName("ec")
            .SetTitle("Electricity Costs")
            .SetDescription("A tool to manage spreadsheet to track Electricity Costs")
            .Build()
            .RunAsync();
}