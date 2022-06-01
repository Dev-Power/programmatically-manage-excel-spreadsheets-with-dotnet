osascript -e 'quit app "Microsoft Excel"'
filename="ElectricityCosts.xlsx"
rm $filename
ec closedxml template new $filename
open "/Applications/Microsoft Excel.app" $filename
