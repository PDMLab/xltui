using System.Globalization;
using ClosedXML.Excel;
using Spectre.Console;
using Spectre.Console.Cli;

namespace XlTui;

public static class Program
{
  public static int Main(string[] args)
  {
    var app = new CommandApp();
    app.Configure(cfg =>
    {
      cfg.SetApplicationName("xltui");
      cfg.AddCommand<RenderCommand>("render")
        .WithDescription("Render an Excel sheet to a console document")
        .WithExample(["render", "--file", "sample.xlsx"])
        .WithExample(["render", "--file", "sample.xlsx", "--json"])
        .WithExample(["render", "--file", "sample.xlsx", "--json", "--columns", "Name,Email"])
        .WithExample(["render", "--file", "sample.xlsx", "--sheet", "People", "--style", "table"])
        .WithExample(
          [
            "render",
            "--file",
            "sample.xlsx",
            "--sheet",
            "People",
            "--style",
            "tree",
            "--group-by",
            "Department",
          ]
        )
        .WithExample(
          ["render", "--file", "sample.xlsx", "--sheet-index", "1", "--columns", "Name,Email,Dept"]
        );
    });

    return app.Run(args);
  }
}

// -------------------------------
// Commands
// -------------------------------
public class RenderSettings : CommandSettings
{
  [CommandOption("--file <PATH>")]
  public string FilePath { get; set; } = "sample.xlsx";

  [CommandOption("--sheet <NAME>")]
  public string? SheetName { get; set; }

  [CommandOption("--sheet-index <N>")]
  public int? SheetIndex { get; set; }

  [CommandOption("--style <table|panel|tree>")]
  public string Style { get; set; } = "table";

  [CommandOption("--group-by <COLUMN>")]
  public string? GroupBy { get; set; }

  [CommandOption("--columns <COLUMNS>")]
  public string? ColumnsCsv { get; set; }

  [CommandOption("--title <TITLE>")]
  public string? Title { get; set; }

  [CommandOption("--max-rows <N>")]
  public int? MaxRows { get; set; }

  [CommandOption("--json")]
  public bool Json { get; set; }
}

public class RenderCommand : Command<RenderSettings>
{
  public override int Execute(CommandContext context, RenderSettings settings)
  {
    // Load workbook/sheet
    if (!File.Exists(settings.FilePath))
    {
      // If JSON output requested, write plain error to stderr without ANSI
      if (settings.Json)
      {
        Console.Error.WriteLine($"File not found: {settings.FilePath}");
      }
      else
      {
        AnsiConsole.MarkupLine($"[red]File not found:[/] {settings.FilePath}");
      }
      return -1;
    }

    using var book = new XLWorkbook(settings.FilePath);
    var ws = ResolveWorksheet(book, settings.SheetName, settings.SheetIndex);
    if (ws is null)
    {
      if (settings.Json)
        Console.Error.WriteLine("Worksheet not found.");
      else
        AnsiConsole.MarkupLine("[red]Worksheet not found.[/]");
      return -1;
    }

    // Read data
    var loader = new ExcelLoader();
    var data = loader.ReadSheet(ws, settings.MaxRows);

    // Column filter
    var columns = settings
      .ColumnsCsv?.Split(
        ',',
        StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries
      )
      .ToArray();

    if (columns is { Length: > 0 })
      data = data with { Headers = [.. data.Headers.Where(h => columns.Contains(h))] };

    // Title
    var title = settings.Title ?? $"{Path.GetFileName(settings.FilePath)} — {ws.Name}";

    // Render or JSON
    if (settings.Json)
    {
      var outObj = new Dictionary<string, List<Dictionary<string, object?>>>(
        StringComparer.OrdinalIgnoreCase
      );
      var rows = data
        .Rows.Select(r =>
        {
          var d = new Dictionary<string, object?>();
          foreach (var h in data.Headers)
          {
            var v = r.GetValueOrDefault(h, "");
            d[h] = string.IsNullOrEmpty(v) ? null : (object)v;
          }
          return d;
        })
        .ToList();
      outObj[data.SheetName] = rows;
      var json = System.Text.Json.JsonSerializer.Serialize(
        outObj,
        new System.Text.Json.JsonSerializerOptions { WriteIndented = true }
      );
      Console.WriteLine(json);
      return 0;
    }

    // If JSON modes are not set, render using Spectre.Console. Otherwise output JSON without ANSI.
    if (!settings.Json)
    {
      var ctx = new RenderContext(title);
      IRenderer renderer = settings.Style.ToLowerInvariant() switch
      {
        "panel" => new PanelRenderer(),
        "tree" => new TreeRenderer(settings.GroupBy),
        _ => new TableRenderer(), // default
      };

      renderer.Render(data, ctx);
    }
    else
    {
      // already handled above
    }

    return 0;
  }

  private static IXLWorksheet? ResolveWorksheet(XLWorkbook book, string? name, int? index)
  {
    if (!string.IsNullOrWhiteSpace(name))
      return book.Worksheets.TryGetWorksheet(name, out var ws) ? ws : null;

    if (index.HasValue && index.Value >= 1 && index.Value <= book.Worksheets.Count)
      return book.Worksheet(index.Value);

    // default: first worksheet
    return book.Worksheet(1);
  }
}

// -------------------------------
// Data model & loader
// -------------------------------
public record SheetData(
  string SheetName,
  List<string> Headers,
  List<Dictionary<string, string>> Rows
);

public class ExcelLoader
{
  // Reads the first used row as header, subsequent rows as data
  public SheetData ReadSheet(IXLWorksheet ws, int? maxRows = null)
  {
    var range = ws.RangeUsed() ?? ws.Range(ws.FirstCell().Address, ws.FirstCell().Address);
    var firstRow = range.FirstRowUsed()?.RowNumber() ?? 1;

    var headerRow = ws.Row(firstRow);
    var headers = headerRow.CellsUsed().Select(c => c.GetString().Trim()).ToList();

    var rows = new List<Dictionary<string, string>>();
    var current = firstRow + 1;
    var end = range.LastRow().RowNumber();

    for (var r = current; r <= end; r++)
    {
      if (maxRows.HasValue && rows.Count >= maxRows.Value)
        break;

      var row = ws.Row(r);
      if (row.IsEmpty())
        continue;

      var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
      for (var i = 0; i < headers.Count; i++)
      {
        var cell = row.Cell(i + 1);
        dict[headers[i]] = CellToString(cell);
      }
      rows.Add(dict);
    }

    return new SheetData(ws.Name, headers, rows);
  }

  private static string CellToString(IXLCell cell)
  {
    if (cell.IsEmpty())
      return string.Empty;

    return cell.DataType switch
    {
      XLDataType.Text => cell.GetString(),
      XLDataType.DateTime => cell.GetDateTime()
        .ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture),
      XLDataType.Number => cell.GetDouble().ToString("0.########", CultureInfo.InvariantCulture),
      XLDataType.Boolean => cell.GetBoolean() ? "true" : "false",
      XLDataType.TimeSpan => cell.GetTimeSpan().ToString("c"),
      _ => cell.GetFormattedString(),
    };
  }
}

// -------------------------------
// Render context & renderers
// -------------------------------
public record RenderContext(string Title);

public interface IRenderer
{
  void Render(SheetData data, RenderContext ctx);
}

// Renders a classic table with Spectre.Console Table
public class TableRenderer : IRenderer
{
  public void Render(SheetData data, RenderContext ctx)
  {
    var table = new Table().RoundedBorder().Expand();
    foreach (var h in data.Headers)
      table.AddColumn(new TableColumn($"[bold]{Escape(h)}[/]").Centered());

    foreach (var row in data.Rows)
      table.AddRow(data.Headers.Select(h => Escape(row.GetValueOrDefault(h, ""))).ToArray());

    var panel = new Panel(table)
      .Header($"[bold]{Escape(ctx.Title)}[/]")
      .Border(BoxBorder.Rounded)
      .Expand();

    AnsiConsole.Write(panel);
  }

  private static string Escape(string s) => Markup.Escape(s);
}

// Renders KPI-ish panels if sheet has 2 columns: Key, Value
public class PanelRenderer : IRenderer
{
  public void Render(SheetData data, RenderContext ctx)
  {
    var grid = new Grid().AddColumn(new GridColumn().NoWrap()).AddColumn(new GridColumn().NoWrap());
    grid.AddRow(
      new Markup($"[bold underline]{Markup.Escape(ctx.Title)}[/]"),
      new Markup($"[dim]{Markup.Escape(data.SheetName)}[/]")
    );
    grid.AddEmptyRow();

    // If sheet is 2 columns, render as Key -> Value panels
    if (data.Headers.Count == 2)
    {
      foreach (var row in data.Rows)
      {
        var key = row.GetValueOrDefault(data.Headers[0], "");
        var val = row.GetValueOrDefault(data.Headers[1], "");
        var p = new Panel($"[bold]{Markup.Escape(val)}[/]")
          .Header(Markup.Escape(key))
          .Border(BoxBorder.Rounded)
          .Expand();
        AnsiConsole.Write(p);
      }
    }
    else
    {
      // Fallback: dump as table
      new TableRenderer().Render(data, ctx);
      return;
    }
  }
}

// Renders a tree grouped by a specified column
public class TreeRenderer : IRenderer
{
  private readonly string? _groupBy;

  public TreeRenderer(string? groupBy)
  {
    _groupBy = groupBy;
  }

  public void Render(SheetData data, RenderContext ctx)
  {
    if (
      string.IsNullOrWhiteSpace(_groupBy)
      || !data.Headers.Contains(_groupBy!, StringComparer.OrdinalIgnoreCase)
    )
    {
      AnsiConsole.MarkupLine("[yellow]No or unknown --group-by column. Falling back to table.[/]");
      new TableRenderer().Render(data, ctx);
      return;
    }

    var root = new Tree($"[bold]{Markup.Escape(ctx.Title)}[/]").Guide(TreeGuide.Ascii);
    var groups = data.Rows.GroupBy(r => r.GetValueOrDefault(_groupBy!, ""));

    foreach (var g in groups.OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
    {
      var node = root.AddNode($"[green]{Markup.Escape(g.Key)}[/]");
      foreach (var row in g)
      {
        // Show the row as a compact line "H1=V1 | H2=V2 | ..."
        var line = string.Join(
          " | ",
          data.Headers.Select(h =>
          {
            var v = row.GetValueOrDefault(h, "");
            return $"[dim]{Markup.Escape(h)}[/]=[white]{Markup.Escape(v)}[/]";
          })
        );
        node.AddNode(new Markup(line));
      }
    }

    AnsiConsole.Write(root);
  }
}
