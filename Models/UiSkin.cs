namespace PptxAvalonia.Models;

/// <summary>Application chrome skins mimicking popular presentation apps.</summary>
public enum UiSkin
{
    LibreOffice = 0,
    GoogleSlides = 1,
    Wps = 2,
    FreeOffice = 3,
    PowerPoint = 4
}

public static class UiSkinInfo
{
    public static string DisplayName(UiSkin skin) => skin switch
    {
        UiSkin.GoogleSlides => "Google 簡報",
        UiSkin.LibreOffice => "LibreOffice Impress",
        UiSkin.Wps => "WPS Presentation",
        UiSkin.FreeOffice => "FreeOffice Presentations",
        UiSkin.PowerPoint => "Microsoft PowerPoint",
        _ => skin.ToString()
    };

    public static string ShortName(UiSkin skin) => skin switch
    {
        UiSkin.GoogleSlides => "Google",
        UiSkin.LibreOffice => "Impress",
        UiSkin.Wps => "WPS",
        UiSkin.FreeOffice => "FreeOffice",
        UiSkin.PowerPoint => "PowerPoint",
        _ => skin.ToString()
    };

    /// <summary>Chrome layout family for the selected skin.</summary>
    public static UiChromeKind ChromeKind(UiSkin skin) => skin switch
    {
        UiSkin.GoogleSlides => UiChromeKind.Google,
        UiSkin.LibreOffice => UiChromeKind.Classic,
        _ => UiChromeKind.Ribbon
    };
}

public enum UiChromeKind
{
    Classic,
    Google,
    Ribbon
}
