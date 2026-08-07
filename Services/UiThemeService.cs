using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using PptxAvalonia.Models;

namespace PptxAvalonia.Services;

/// <summary>Applies presentation-app color palettes to application resources at runtime.</summary>
public static class UiThemeService
{
    public static void Apply(UiSkin skin)
    {
        var app = Application.Current;
        if (app is null) return;

        var p = PaletteFor(skin);
        var res = app.Resources;

        SetBrush(res, "UiAccentBrush", p.Accent);
        SetBrush(res, "UiAccentDarkBrush", p.AccentDark);
        SetBrush(res, "UiAccentSoftBrush", p.AccentSoft);
        SetBrush(res, "UiSelectBrush", p.Select);
        SetBrush(res, "UiSelectBgBrush", p.SelectBg);
        SetBrush(res, "UiInkBrush", p.Ink);
        SetBrush(res, "UiMutedBrush", p.Muted);
        SetBrush(res, "UiBorderBrush", p.Border);
        SetBrush(res, "UiBorderLightBrush", p.BorderLight);
        SetBrush(res, "UiSurfaceBrush", p.Surface);
        SetBrush(res, "UiChromeBrush", p.Chrome);
        SetBrush(res, "UiToolbarBrush", p.Toolbar);
        SetBrush(res, "UiCanvasBrush", p.Canvas);
        SetBrush(res, "UiSidebarBrush", p.Sidebar);
        SetBrush(res, "UiSidebarHeaderBrush", p.SidebarHeader);
        SetBrush(res, "UiStatusBrush", p.Status);
        SetBrush(res, "UiFindBrush", p.Find);
        SetBrush(res, "UiHoverBrush", p.Hover);
        SetBrush(res, "UiPressedBrush", p.Pressed);
        SetBrush(res, "UiRibbonTabBarBrush", p.RibbonTabBar);
        SetBrush(res, "UiRibbonBodyBrush", p.RibbonBody);
        SetBrush(res, "UiQatBrush", p.Qat);
        SetBrush(res, "UiPrimaryFgBrush", p.PrimaryFg);
        SetBrush(res, "UiOnAccentBrush", p.OnAccent);

        // Keep Lo* aliases so any leftover StaticResource bindings still resolve.
        SetBrush(res, "LoGreenBrush", p.Accent);
        SetBrush(res, "LoGreenDarkBrush", p.AccentDark);
        SetBrush(res, "LoGreenSoftBrush", p.AccentSoft);
        SetBrush(res, "LoSelectBrush", p.Select);
        SetBrush(res, "LoSelectBgBrush", p.SelectBg);
        SetBrush(res, "LoInkBrush", p.Ink);
        SetBrush(res, "LoMutedBrush", p.Muted);
        SetBrush(res, "LoBorderBrush", p.Border);
        SetBrush(res, "LoBorderLightBrush", p.BorderLight);
        SetBrush(res, "LoSurfaceBrush", p.Surface);
        SetBrush(res, "LoChromeBrush", p.Chrome);
        SetBrush(res, "LoToolbarBrush", p.Toolbar);
        SetBrush(res, "LoCanvasBrush", p.Canvas);
        SetBrush(res, "LoSidebarBrush", p.Sidebar);
        SetBrush(res, "LoSidebarHeaderBrush", p.SidebarHeader);
        SetBrush(res, "LoStatusBrush", p.Status);
        SetBrush(res, "LoFindBrush", p.Find);
        SetBrush(res, "LoHoverBrush", p.Hover);
        SetBrush(res, "LoPressedBrush", p.Pressed);
    }

    private static void SetBrush(IResourceDictionary res, string key, Color color)
    {
        if (res.TryGetResource(key, theme: null, out var existing) && existing is SolidColorBrush brush)
        {
            brush.Color = color;
            return;
        }

        res[key] = new SolidColorBrush(color);
    }

    private static Palette PaletteFor(UiSkin skin) => skin switch
    {
        UiSkin.GoogleSlides => new Palette(
            Accent: Color.Parse("#1A73E8"),
            AccentDark: Color.Parse("#1557B0"),
            AccentSoft: Color.Parse("#E8F0FE"),
            Select: Color.Parse("#1A73E8"),
            SelectBg: Color.Parse("#D2E3FC"),
            Ink: Color.Parse("#202124"),
            Muted: Color.Parse("#5F6368"),
            Border: Color.Parse("#DADCE0"),
            BorderLight: Color.Parse("#E8EAED"),
            Surface: Colors.White,
            Chrome: Colors.White,
            Toolbar: Color.Parse("#F8F9FA"),
            Canvas: Color.Parse("#F1F3F4"),
            Sidebar: Color.Parse("#F8F9FA"),
            SidebarHeader: Colors.White,
            Status: Color.Parse("#F8F9FA"),
            Find: Color.Parse("#1A73E8"),
            Hover: Color.Parse("#F1F3F4"),
            Pressed: Color.Parse("#E8EAED"),
            RibbonTabBar: Color.Parse("#1A73E8"),
            RibbonBody: Color.Parse("#F8F9FA"),
            Qat: Colors.White,
            PrimaryFg: Color.Parse("#1A73E8"),
            OnAccent: Colors.White),

        UiSkin.LibreOffice => new Palette(
            Accent: Color.Parse("#18A303"),
            AccentDark: Color.Parse("#0E7A00"),
            AccentSoft: Color.Parse("#E8F5E4"),
            Select: Color.Parse("#308CC6"),
            SelectBg: Color.Parse("#CDE8F6"),
            Ink: Color.Parse("#1A1A1A"),
            Muted: Color.Parse("#4D4D4D"),
            Border: Color.Parse("#A8A8A8"),
            BorderLight: Color.Parse("#C8C8C8"),
            Surface: Colors.White,
            Chrome: Color.Parse("#EDEDED"),
            Toolbar: Color.Parse("#F5F5F5"),
            Canvas: Color.Parse("#808080"),
            Sidebar: Color.Parse("#F0F0F0"),
            SidebarHeader: Color.Parse("#E0E0E0"),
            Status: Color.Parse("#E8E8E8"),
            Find: Color.Parse("#18A303"),
            Hover: Color.Parse("#D8D8D8"),
            Pressed: Color.Parse("#C0C0C0"),
            RibbonTabBar: Color.Parse("#18A303"),
            RibbonBody: Color.Parse("#F5F5F5"),
            Qat: Color.Parse("#EDEDED"),
            PrimaryFg: Color.Parse("#0E7A00"),
            OnAccent: Colors.White),

        UiSkin.Wps => new Palette(
            Accent: Color.Parse("#C72A1C"),
            AccentDark: Color.Parse("#A01F14"),
            AccentSoft: Color.Parse("#FDECEA"),
            Select: Color.Parse("#C72A1C"),
            SelectBg: Color.Parse("#FADBD8"),
            Ink: Color.Parse("#222222"),
            Muted: Color.Parse("#666666"),
            Border: Color.Parse("#D0D0D0"),
            BorderLight: Color.Parse("#E5E5E5"),
            Surface: Colors.White,
            Chrome: Color.Parse("#F7F7F7"),
            Toolbar: Colors.White,
            Canvas: Color.Parse("#6B6B6B"),
            Sidebar: Color.Parse("#F5F5F5"),
            SidebarHeader: Color.Parse("#EEEEEE"),
            Status: Color.Parse("#F0F0F0"),
            Find: Color.Parse("#C72A1C"),
            Hover: Color.Parse("#F0F0F0"),
            Pressed: Color.Parse("#E0E0E0"),
            RibbonTabBar: Color.Parse("#C72A1C"),
            RibbonBody: Color.Parse("#FFF8F7"),
            Qat: Color.Parse("#F3F3F3"),
            PrimaryFg: Color.Parse("#C72A1C"),
            OnAccent: Colors.White),

        UiSkin.FreeOffice => new Palette(
            Accent: Color.Parse("#0072C6"),
            AccentDark: Color.Parse("#005A9E"),
            AccentSoft: Color.Parse("#E8F1FA"),
            Select: Color.Parse("#0072C6"),
            SelectBg: Color.Parse("#D0E4F7"),
            Ink: Color.Parse("#222222"),
            Muted: Color.Parse("#666666"),
            Border: Color.Parse("#B4C7D9"),
            BorderLight: Color.Parse("#D0DCE8"),
            Surface: Colors.White,
            Chrome: Color.Parse("#F3F3F3"),
            Toolbar: Color.Parse("#E8F1FA"),
            Canvas: Color.Parse("#7A7A7A"),
            Sidebar: Color.Parse("#F0F0F0"),
            SidebarHeader: Color.Parse("#E6E6E6"),
            Status: Color.Parse("#F0F0F0"),
            Find: Color.Parse("#0072C6"),
            Hover: Color.Parse("#D0E4F7"),
            Pressed: Color.Parse("#B8D4F0"),
            RibbonTabBar: Color.Parse("#0072C6"),
            RibbonBody: Color.Parse("#E8F1FA"),
            Qat: Color.Parse("#F3F3F3"),
            PrimaryFg: Color.Parse("#005A9E"),
            OnAccent: Colors.White),

        UiSkin.PowerPoint => new Palette(
            Accent: Color.Parse("#C43E1C"),
            AccentDark: Color.Parse("#A43316"),
            AccentSoft: Color.Parse("#FCE9E4"),
            Select: Color.Parse("#2B579A"),
            SelectBg: Color.Parse("#D6E3F5"),
            Ink: Color.Parse("#252423"),
            Muted: Color.Parse("#605E5C"),
            Border: Color.Parse("#C8C6C4"),
            BorderLight: Color.Parse("#E1DFDD"),
            Surface: Colors.White,
            Chrome: Color.Parse("#F3F2F1"),
            Toolbar: Color.Parse("#FAF9F8"),
            Canvas: Color.Parse("#E6E6E6"),
            Sidebar: Color.Parse("#F3F2F1"),
            SidebarHeader: Color.Parse("#EDEBE9"),
            Status: Color.Parse("#F3F2F1"),
            Find: Color.Parse("#C43E1C"),
            Hover: Color.Parse("#E1DFDD"),
            Pressed: Color.Parse("#C8C6C4"),
            RibbonTabBar: Color.Parse("#C43E1C"),
            RibbonBody: Color.Parse("#FAF9F8"),
            Qat: Color.Parse("#F3F2F1"),
            PrimaryFg: Color.Parse("#C43E1C"),
            OnAccent: Colors.White),

        _ => PaletteFor(UiSkin.LibreOffice)
    };

    private readonly record struct Palette(
        Color Accent,
        Color AccentDark,
        Color AccentSoft,
        Color Select,
        Color SelectBg,
        Color Ink,
        Color Muted,
        Color Border,
        Color BorderLight,
        Color Surface,
        Color Chrome,
        Color Toolbar,
        Color Canvas,
        Color Sidebar,
        Color SidebarHeader,
        Color Status,
        Color Find,
        Color Hover,
        Color Pressed,
        Color RibbonTabBar,
        Color RibbonBody,
        Color Qat,
        Color PrimaryFg,
        Color OnAccent);
}
