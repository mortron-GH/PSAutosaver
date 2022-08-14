using System;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.ComponentModel;
using MahApps.Metro.Controls;

namespace PSAutosaver
{
	public partial class MainWindow : MetroWindow
	{

		private const string scriptPath = "PSAutosave.VBS";
		private const string savePath = "Settings.json";

		private bool started;

		private SaveData settings;

		public MainWindow()
		{
			InitializeComponent();

			Initialize();
		}

		private void Initialize()
		{
			settings = LoadSettings();

			StartStopButton.Click += HandleStartStopButtonClicked;
			ShowDialogToggle.Click += HandleShowDialogToggleChecked;
			IntervalSlider.ValueChanged += HandleSliderValueChanged;
			InfoLabel.MouseUp += HandleInfoLabelClicked;

			CreateScript();

			UpdateSliderLabelView();
			UpdateSliderAndCheckBox();
			InitInfoText();
		}

		protected override void OnClosing(CancelEventArgs e)
		{
			base.OnClosing(e);

			SaveSettings();
		}

		private void InitInfoText()
		{
			byte[] data = Convert.FromBase64String("bWFkZSBmb3Ig8J+MuEFpcmVuIHdpdGgg4pml");
			InfoTextBlock.Text = Encoding.UTF8.GetString(data);
		}

		private void HandleInfoLabelClicked(object o, System.Windows.Input.MouseButtonEventArgs e)
		{
			var visibility = InfoRectangle.IsVisible ? Visibility.Hidden : Visibility.Visible;
			InfoRectangle.Visibility = visibility;
			InfoTextBlock.Visibility = visibility;

			switch (visibility)
			{
				case Visibility.Visible:
					this.MouseUp += HandleWindowMouseClick;
					System.Windows.Input.Keyboard.ClearFocus();
					break;

				default:
					this.MouseUp -= HandleWindowMouseClick;
					break;
			}
		}

		private void HandleWindowMouseClick(object o, System.Windows.Input.MouseButtonEventArgs e)
		{
			HandleInfoLabelClicked(null, null);
		}

		private void HandleSliderValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
		{
			var time = MathF.Truncate(600.0f * (float)(IntervalSlider.Value / IntervalSlider.Maximum));
			settings.DelaySeconds = (int)MathF.Max(5.0f, time);

			UpdateSliderLabelView();
		}

		private void UpdateSliderLabelView()
		{
			TimeSpan ts = TimeSpan.FromSeconds(settings.DelaySeconds);
			IntervalSliderLabel.Content = "Save every " + ts.ToString(@"mm\:ss");
		}

		private void UpdateSliderAndCheckBox()
		{
			IntervalSlider.Value = IntervalSlider.Maximum * (settings.DelaySeconds / 600.0f);
			ShowDialogToggle.IsChecked = settings.ShowSaveDialog;
		}

		private void HandleShowDialogToggleChecked(object o, RoutedEventArgs e)
		{
			var dialog = (bool)ShowDialogToggle.IsChecked;

			if (dialog != settings.ShowSaveDialog)
			{
				settings.ShowSaveDialog = dialog;
				CreateScript();
			}
		}

		private async void HandleStartStopButtonClicked(object o , RoutedEventArgs e)
		{
			started = !started;

			StartStopButton.Content = started ? "Stop" : "Start";

			await MainTask();
		}

		private async Task MainTask()
		{
			Process process = null;

			while (true)
			{
				await Task.Delay(settings.DelaySeconds * 1000);

				process?.Close();
				process?.Dispose();

				if (!started)
					break;

				process = SetupProcess();
				process.Start();

				await process.WaitForExitAsync();
			}
		}

		private Process SetupProcess()
		{
			var path = AppDomain.CurrentDomain.BaseDirectory;
			var p = new Process();
			p.StartInfo.FileName = @"cscript";
			p.StartInfo.WorkingDirectory = path;
			p.StartInfo.Arguments = "//B //Nologo " + Path.Combine(path, scriptPath);
			p.StartInfo.CreateNoWindow = true;

			return p;
		}

		private void CreateScript()
		{
			int dialogMode = settings.ShowSaveDialog ? 1 : 3;

			var data = string.Format("DIM objApp" +
				"\nSET objApp = CreateObject(\"Photoshop.Application\")" +
				"\nDIM dialogMode" +
				"\ndialogMode = {0}" +
				"\nDIM id5809" +
				"\nid5809 = objApp.CharIDToTypeID(\"save\")" +
				"\nCall objApp.ExecuteAction(id5809, , dialogMode)",
				dialogMode);

			StreamWriter file = File.CreateText(scriptPath);

			file.WriteLine(data);
			file.Close();
		}

		private SaveData LoadSettings()
		{
			if (File.Exists(savePath))
			{
				var data = File.ReadAllText(savePath);
				return JsonSerializer.Deserialize<SaveData>(data);
			}
			else
			{
				return SaveData.Default;
			}
		}

		private void SaveSettings()
		{
			var data = JsonSerializer.Serialize(settings);

			var file = File.CreateText(savePath);
			file.WriteLine(data);
			file.Close();
		}
	}

	public class SaveData
	{
		public int DelaySeconds { get; set; }
		public bool ShowSaveDialog { get; set; }

		public static SaveData Default => new SaveData() { DelaySeconds = 5, ShowSaveDialog = false };
	}
}
