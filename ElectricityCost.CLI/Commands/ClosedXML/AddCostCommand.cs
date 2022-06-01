using CliFx;
using CliFx.Attributes;
using CliFx.Infrastructure;
using ElectricityCost.CLI.Commands.Core;

namespace ElectricityCost.CLI.Commands.ClosedXML;

[Command("closedxml add")]
public class AddCostCommand : ICommand
{
    [CommandOption("path", 'p', Description = "The path to the Excel spreadsheet")]
    public string Path { get; init; }
    
    [CommandOption("day", 'd', Description = "The day of the month")]
    public decimal DayOfMonth { get; init; }
    
    [CommandOption("cost", 'c', Description = "The cost of electricity for the day")]
    public decimal Cost { get; init; }
    
    public ValueTask ExecuteAsync(IConsole console)
    {
        var templateManager = new TemplateManager(Path);
        templateManager.AddCost(DayOfMonth, Cost);
        return default;
    }
}