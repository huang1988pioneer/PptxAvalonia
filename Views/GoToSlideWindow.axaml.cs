using Avalonia.Controls;
using Avalonia.Interactivity;

namespace PptxAvalonia.Views;

public partial class GoToSlideWindow : Window
{
    public int? ResultIndex { get; private set; }

    public GoToSlideWindow()
    {
        InitializeComponent();
        OkButton.Click += OnOk;
        CancelButton.Click += (_, _) => Close();
        Opened += (_, _) => NumberBox.Focus();
    }

    public void Configure(int currentOneBased, int max)
    {
        PromptText.Text = $"輸入投影片編號（1–{max}）：";
        NumberBox.Text = currentOneBased.ToString();
        NumberBox.SelectAll();
    }

    private void OnOk(object? sender, RoutedEventArgs e)
    {
        if (int.TryParse(NumberBox.Text?.Trim(), out var n) && n >= 1)
        {
            ResultIndex = n - 1;
            Close(ResultIndex);
        }
        else
        {
            NumberBox.Focus();
            NumberBox.SelectAll();
        }
    }
}
