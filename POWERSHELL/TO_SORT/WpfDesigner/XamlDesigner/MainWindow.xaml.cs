using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ICSharpCode.XamlDesigner.Configuration;
using System.ComponentModel;
using Microsoft.Win32;
using System.IO;
using ICSharpCode.WpfDesign.Designer;
using AvalonDock.Layout.Serialization;

namespace ICSharpCode.XamlDesigner
{
	public partial class MainWindow
	{
		public MainWindow()
		{
			Instance = this;
			DataContext = Shell.Instance;
			RenameCommands();
			BasicMetadata.Register();

			InitializeComponent();

			Shell.Instance.PropertyGrid = uxPropertyGridView.PropertyGrid;
			//AvalonDockWorkaround();
			RouteDesignSurfaceCommands();

			this.AddCommandHandler(RefreshCommand, Shell.Instance.Refresh, Shell.Instance.CanRefresh);

			LoadSettings();
			ProcessPaths(App.Args);

			ApplicationCommands.New.Execute(null, this);
		}

		public static MainWindow Instance;

		OpenFileDialog openFileDialog;
		SaveFileDialog saveFileDialog;

		protected override void OnDragEnter(DragEventArgs e)
		{
			ProcessDrag(e);
		}

		protected override void OnDragOver(DragEventArgs e)
		{
			ProcessDrag(e);
		}

		protected override void OnDrop(DragEventArgs e)
		{
			ProcessPaths(e.Data.Paths());
		}

		protected override void OnClosing(CancelEventArgs e)
		{
			if (Shell.Instance.PrepareExit())
			{
				SaveSettings();
			}
			else
			{
				e.Cancel = true;
			}
			base.OnClosing(e);
		}

		void RecentFiles_Click(object sender, RoutedEventArgs e)
		{
			var path = (string)(e.OriginalSource as MenuItem).Header;
			Shell.Instance.Open(path);
		}

		void ProcessDrag(DragEventArgs e)
		{
			e.Effects = DragDropEffects.None;
			e.Handled = true;

			foreach (var path in e.Data.Paths())
			{
				if (path.EndsWith(".dll", StringComparison.InvariantCultureIgnoreCase) ||
					path.EndsWith(".exe", StringComparison.InvariantCultureIgnoreCase))
				{
					e.Effects = DragDropEffects.Copy;
					break;
				}
				else if (path.EndsWith(".xaml", StringComparison.InvariantCultureIgnoreCase))
				{
					e.Effects = DragDropEffects.Copy;
					break;
				}
			}
		}

		void ProcessPaths(IEnumerable<string> paths)
		{
			foreach (var path in paths)
			{
				if (path.EndsWith(".dll", StringComparison.InvariantCultureIgnoreCase) ||
					path.EndsWith(".exe", StringComparison.InvariantCultureIgnoreCase))
				{
					Toolbox.Instance.AddAssembly(path);
				}
				else if (path.EndsWith(".xaml", StringComparison.InvariantCultureIgnoreCase))
				{
					Shell.Instance.Open(path);
				}
			}
		}

		public string AskOpenFileName()
		{
			if (openFileDialog == null)
			{
				openFileDialog = new OpenFileDialog();
				openFileDialog.Filter = "Xaml Documents (*.xaml)|*.xaml";
			}
			if ((bool)openFileDialog.ShowDialog())
			{
				return openFileDialog.FileName;
			}
			return null;
		}

		public string AskSaveFileName(string initName)
		{
			if (saveFileDialog == null)
			{
				saveFileDialog = new SaveFileDialog();
				saveFileDialog.Filter = "Xaml Documents (*.xaml)|*.xaml";
			}
			saveFileDialog.FileName = initName;
			if ((bool)saveFileDialog.ShowDialog())
			{
				return saveFileDialog.FileName;
			}
			return null;
		}

		void LoadSettings()
		{
			WindowState = Settings.Default.MainWindowState;

			Rect r = Settings.Default.MainWindowRect;
			if (r != new Rect())
			{
				Left = r.Left;
				Top = r.Top;
				Width = r.Width;
				Height = r.Height;
			}

			uxDockingManager.Loaded += delegate
			{
				if (Settings.Default.AvalonDockLayout != null)
				{
					XmlLayoutSerializer layoutSerializer = new XmlLayoutSerializer(uxDockingManager);
					using (var reader = new StringReader(Settings.Default.AvalonDockLayout)) {
						layoutSerializer.Deserialize(reader);
					}
				}
			};
		}

		void SaveSettings()
		{
			Settings.Default.MainWindowState = WindowState;
			if (WindowState == WindowState.Normal) {
				Settings.Default.MainWindowRect = new Rect(Left, Top, Width, Height);
			}

			XmlLayoutSerializer layoutSerializer = new XmlLayoutSerializer(uxDockingManager);
			using (var writer = new StringWriter()) {
				layoutSerializer.Serialize(writer);
				Settings.Default.AvalonDockLayout = writer.ToString();
			}

			Shell.Instance.SaveSettings();
		}
	}
}
