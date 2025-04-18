using yuseok.kim.dw2docs.Common.Constants;
using yuseok.kim.dw2docs.Common.Extensions;
using System.Text;

namespace yuseok.kim.dw2docs.Common.Utils.CodeTable;

public class CodeTableTools
{
    public static bool GetValueMap(string codetableString, out IDictionary<string, string>? codetable, out string? error)
    {
        error = null;
        var dict = new Dictionary<string, string>();

        try
        {
            string[] tokens = codetableString.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string token in tokens)
            {
                var items = token.Split(new[] { '\t' });
                dict[items[1]] = items[0];
            }

        }
        catch (Exception e)
        {
            dict = null;
            error = e.Message;
        }
        codetable = dict;
        return error is null;
    }

    public static string[] GetDisplayValues(IDictionary<string, string> codemap) => codemap.Values.ToArray();

    public static string[] GetDataValues(IDictionary<string, string> codemap) => codemap.Keys.ToArray();

    public static string? BuildString(IDictionary<string, string> codetable, string? value, bool left, int columns)
    {
        var sb = new StringBuilder();

        int i = 0;
        foreach (var (data, display) in codetable)
        {
            if (!left)
            {
                sb.Append($"{GetButton(codetable, value, display)} ");
            }
            sb.Append($"{display}");

            if (left)
            {
                sb.Append($" {GetButton(codetable, value, display)}");
            }

            if ((i + 1) % columns == 0)
                sb.AppendLine();
            else
                sb.Append("    ");

            ++i;
        }

        sb.Length -= 2;

        return sb.ToString();
    }

    private static string? GetButton(IDictionary<string, string> codetable, string? value, string expected)
            => (value is null
            ? RendererConstants.RadioButtonUnselected
            : codetable.ContainsKey(value) && codetable[value] == expected
                ? RendererConstants.RadioButtonSelected
                : RendererConstants.RadioButtonUnselected).ToString();
}
