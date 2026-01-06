using System;
using System.ComponentModel;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace QuickOptionImportExport;

/// <summary>
/// Row wrapper around the underlying JSON object.
/// Only a subset of fields are shown in the UI, but the JSON node is preserved for round-tripping.
/// </summary>
public sealed class OptionRow : INotifyPropertyChanged
{
    /// <summary>
    /// Factory used to create JSON for new rows ("Add new row").
    /// MainForm sets this after importing a file (clone-first-row behavior).
    /// </summary>
    public static Func<JsonObject> TemplateFactory { get; set; } = CreateDefaultTemplate;

    /// <summary>
    /// Creates a minimal default option object. Used when starting a new file.
    /// </summary>
    public static JsonObject CreateDefaultTemplate() => DefaultTemplateFactory();

    public JsonObject Node { get; }

    public OptionRow() : this(TemplateFactory()) { }

    public OptionRow(JsonObject node)
    {
        Node = node ?? throw new ArgumentNullException(nameof(node));
        EnsureRequiredShape(Node);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void Raise(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    // === Top-level fields ===

    public int OptionNumber
    {
        get => ReadInt(Node, "OptionNumber");
        set { WriteInt(Node, "OptionNumber", value); Raise(nameof(OptionNumber)); }
    }

    public int Sequence
    {
        get => ReadInt(Node, "OptionSequence");
        set { WriteInt(Node, "OptionSequence", value); Raise(nameof(Sequence)); }
    }

    // === Information fields ===

    private JsonObject Info => (JsonObject)Node["Information"]!;

    public string TextValue
    {
        get => ReadString(Info, "Text");
        set { WriteString(Info, "Text", value); Raise(nameof(TextValue)); }
    }

    public string AsciiValue
    {
        get => ReadString(Info, "AsciiFlag");
        set { WriteString(Info, "AsciiFlag", value); Raise(nameof(AsciiValue)); }
    }

    public bool BooleanValue
    {
        get => ReadBool(Info, "Boolean");
        set { WriteBool(Info, "Boolean", value); Raise(nameof(BooleanValue)); }
    }

    public double NumericValue
    {
        get => ReadDouble(Info, "Numeric");
        set { WriteDouble(Info, "Numeric", value); Raise(nameof(NumericValue)); }
    }

    public long LongValue
    {
        get => ReadLong(Info, "Long");
        set { WriteLong(Info, "Long", value); Raise(nameof(LongValue)); }
    }

    /// <summary>
    /// Date part shown in grid. Stored as string for simple DataGridView binding.
    /// Expected format: yyyy-MM-dd (empty allowed).
    /// </summary>
    public string DateValue
    {
        get
        {
            var dt = ReadDateTime(Info, "DateTime");
            return dt.Year <= 1900 ? "" : dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }
        set
        {
            var current = ReadDateTime(Info, "DateTime");
            if (string.IsNullOrWhiteSpace(value))
            {
                // Keep time component if it exists; default date to 1900-01-01.
                var reset = new DateTime(1900, 1, 1, current.Hour, current.Minute, current.Second);
                WriteDateTime(Info, "DateTime", reset);
            }
            else
            {
                var date = DateOnly.ParseExact(value.Trim(), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                var updated = new DateTime(date.Year, date.Month, date.Day, current.Hour, current.Minute, current.Second);
                WriteDateTime(Info, "DateTime", updated);
            }
            Raise(nameof(DateValue));
        }
    }

    /// <summary>
    /// Time part shown in grid. Stored as string for simple DataGridView binding.
    /// Expected format: HH:mm:ss or HH:mm (empty allowed).
    /// </summary>
    public string TimeValue
    {
        get
        {
            var dt = ReadDateTime(Info, "DateTime");
            if (dt.Year <= 1900 && dt.Hour == 0 && dt.Minute == 0 && dt.Second == 0) return "";
            return dt.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
        }
        set
        {
            var current = ReadDateTime(Info, "DateTime");
            if (string.IsNullOrWhiteSpace(value))
            {
                var reset = new DateTime(current.Year, current.Month, current.Day, 0, 0, 0);
                WriteDateTime(Info, "DateTime", reset);
            }
            else
            {
                var t = ParseTime(value.Trim());
                var updated = new DateTime(current.Year, current.Month, current.Day, t.Hours, t.Minutes, t.Seconds);
                WriteDateTime(Info, "DateTime", updated);
            }
            Raise(nameof(TimeValue));
        }
    }

    public static TimeSpan ParseTime(string s)
    {
        // Accept HH:mm or HH:mm:ss
        if (TimeSpan.TryParseExact(s, new[] { "hh\\:mm", "hh\\:mm\\:ss", "HH\\:mm", "HH\\:mm\\:ss" },
                CultureInfo.InvariantCulture, out var ts))
            return ts;

        // Last resort: TimeSpan.Parse handles many patterns.
        return TimeSpan.Parse(s, CultureInfo.InvariantCulture);
    }

    // === Helpers ===

    private static void EnsureRequiredShape(JsonObject node)
    {
        node["OptionNumber"] ??= 0;
        node["OptionSequence"] ??= 0;

        if (node["Information"] is not JsonObject info)
        {
            info = new JsonObject();
            node["Information"] = info;
        }

        info["Text"] ??= "";
        info["AsciiFlag"] ??= "";
        info["Boolean"] ??= false;
        info["Long"] ??= 0;
        info["Numeric"] ??= 0.0;
        info["DateTime"] ??= "1900-01-01T00:00:00";
    }

    private static JsonObject DefaultTemplateFactory()
    {
        // A minimal-but-compatible template. If the user imported an existing file, MainForm overrides
        // TemplateFactory with a deep clone of the first item for near-perfect formatting.
        return new JsonObject
        {
            ["_IsModified"] = false,
            ["OptionNumber"] = 0,
            ["OptionSequence"] = 0,
            ["SystemAudit"] = new JsonObject
            {
                ["LastChange"] = new JsonObject
                {
                    ["_DateTime"] = DateTime.Now.ToString("O"),
                    ["_DateTimeReadCount"] = 0,
                    ["_DateTimeWriteCount"] = 0,
                    ["_TerminalCode"] = "",
                    ["_TerminalCodeReadCount"] = 0,
                    ["_TerminalCodeWriteCount"] = 0,
                    ["UserCD"] = new JsonObject { ["Information"] = new JsonObject { ["Username"] = "" } },
                    ["ProgramCD"] = new JsonObject { ["Information"] = new JsonObject { ["_ShortName"] = "", ["_ShortNameReadCount"] = 0, ["_ShortNameWriteCount"] = 0, ["ShortName"] = "" } },
                    ["DateTime"] = DateTime.Now.ToString("O"),
                    ["TerminalCode"] = ""
                },
                ["CompareResults"] = new JsonObject { ["m_MaxCapacity"] = int.MaxValue, ["Capacity"] = 16, ["m_StringValue"] = "", ["m_currentThread"] = 0 },
                ["Db"] = null,
                ["GabProperties"] = null
            },
            ["Obsolete"] = new JsonObject
            {
                ["LastReadBy"] = "",
                ["Filler"] = "",
                ["LastReadDate"] = new JsonObject { ["DateTime"] = "1900-01-01T00:00:00" }
            },
            ["Information"] = new JsonObject
            {
                ["Text"] = "",
                ["AsciiFlag"] = "",
                ["Boolean"] = false,
                ["Long"] = 0,
                ["Numeric"] = 0.0,
                ["DateTime"] = "1900-01-01T00:00:00"
            },
            ["GabProperties"] = null,
            ["IsModified"] = false,
            ["LastModeRead"] = 0,
            ["ChildObject"] = null,
            ["ContextStatus"] = 0,
            ["ModePropertiesRead"] = null,
            ["ConnectionIndex"] = 0,
            ["Locked"] = false,
            ["ProviderStatus"] = 0,
            ["ProviderStatusDescription"] = "",
            ["CompareResults"] = new JsonObject { ["m_MaxCapacity"] = int.MaxValue, ["Capacity"] = 16, ["m_StringValue"] = "", ["m_currentThread"] = 0 },
            ["Db"] = new JsonObject
            {
                ["DatabaseEngineType"] = 0,
                ["Bt"] = new JsonObject
                {
                    ["CompanyCode"] = "",
                    ["OverrideLock"] = false,
                    ["SuppressBTErrorUI"] = false
                },
                ["CompanyCode"] = "",
                ["ServerName"] = "",
                ["IsDisposed"] = false
            }
        };
    }

    private static int ReadInt(JsonObject obj, string key)
        => TryGet(obj, key, out var n) && n is JsonValue v && v.TryGetValue<int>(out var i) ? i : 0;

    private static void WriteInt(JsonObject obj, string key, int value) => obj[key] = value;

    private static long ReadLong(JsonObject obj, string key)
        => TryGet(obj, key, out var n) && n is JsonValue v && v.TryGetValue<long>(out var l) ? l : 0L;

    private static void WriteLong(JsonObject obj, string key, long value) => obj[key] = value;

    private static double ReadDouble(JsonObject obj, string key)
        => TryGet(obj, key, out var n) && n is JsonValue v && v.TryGetValue<double>(out var d) ? d : 0.0;

    private static void WriteDouble(JsonObject obj, string key, double value) => obj[key] = value;

    private static bool ReadBool(JsonObject obj, string key)
        => TryGet(obj, key, out var n) && n is JsonValue v && v.TryGetValue<bool>(out var b) ? b : false;

    private static void WriteBool(JsonObject obj, string key, bool value) => obj[key] = value;

    private static string ReadString(JsonObject obj, string key)
        => TryGet(obj, key, out var n) && n is JsonValue v && v.TryGetValue<string>(out var s) ? s ?? "" : "";

    private static void WriteString(JsonObject obj, string key, string? value) => obj[key] = value ?? "";

    private static DateTime ReadDateTime(JsonObject obj, string key)
    {
        if (!TryGet(obj, key, out var n) || n is null) return new DateTime(1900, 1, 1);

        if (n is JsonValue v)
        {
            if (v.TryGetValue<DateTime>(out var dt)) return dt;
            if (v.TryGetValue<string>(out var s) && DateTime.TryParse(s, null, DateTimeStyles.RoundtripKind, out var parsed))
                return parsed;
        }
        return new DateTime(1900, 1, 1);
    }

    private static void WriteDateTime(JsonObject obj, string key, DateTime dt) => obj[key] = dt.ToString("O");

    private static bool TryGet(JsonObject obj, string key, out JsonNode? node)
        => obj.TryGetPropertyValue(key, out node);
}
