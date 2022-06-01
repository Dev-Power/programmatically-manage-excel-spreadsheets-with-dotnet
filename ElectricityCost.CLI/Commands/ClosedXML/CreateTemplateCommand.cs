using CliFx;
using CliFx.Attributes;
using CliFx.Infrastructure;
using ElectricityCost.CLI.Commands.Core;

namespace ElectricityCost.CLI.Commands.ClosedXML;

[Command("closedxml template new")]
public class NewTemplateCommand : ICommand
{
    [CommandParameter(0, Description = "The path to the template that will be created")]
    public string Path { get; init; }
    
    public ValueTask ExecuteAsync(IConsole console)
    {
        var templateManager = new TemplateManager(Path);
        templateManager.CreateExceltemplate();

        return default;
    }
}