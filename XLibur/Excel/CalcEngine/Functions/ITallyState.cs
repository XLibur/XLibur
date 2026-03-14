namespace XLibur.Excel.CalcEngine.Functions;

internal interface ITallyState<out TState>
{
    TState Tally(double number);
}
