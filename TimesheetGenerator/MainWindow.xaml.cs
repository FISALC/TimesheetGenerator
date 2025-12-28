using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Media;

namespace TimesheetGenerator
{
    public partial class MainWindow : Window
    {
        private readonly TimesheetService _service;

        public MainWindow()
        {
            InitializeComponent();
            _service = new TimesheetService();
            txtYear.Text = DateTime.Now.Year.ToString();
            cmbMonth.SelectedIndex = DateTime.Now.Month - 1;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                txtStatus.Text = "Generating...";
                txtStatus.Foreground = Brushes.Blue;

                if (!int.TryParse(txtYear.Text, out int year))
                {
                    MessageBox.Show("Invalid Year", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                int month = cmbMonth.SelectedIndex + 1;
                
                string approverName = "Salem Jaber Al-Marri";
                string approverRole = "Head of E-Portal Development";

                if (rbApprover2.IsChecked == true)
                {
                    approverName = "Shaikha Abdulrahman AlKhulaifi";
                    approverRole = "Digital Government Systems Expert";
                }

                string filePath = _service.GenerateTimesheet(year, month, approverName, approverRole);
                
                txtStatus.Text = "Success! Saved to Downloads.";
                txtStatus.Foreground = Brushes.Green;

                var result = MessageBox.Show($"Timesheet generated successfully!\n\nLocation: {filePath}\n\nDo you want to open the folder?", 
                                             "Success", MessageBoxButton.YesNo, MessageBoxImage.Information);

                if (result == MessageBoxResult.Yes)
                {
                    Process.Start("explorer.exe", $"/select,\"{filePath}\"");
                }
            }
            catch (Exception ex)
            {
                txtStatus.Text = "Error occurred.";
                txtStatus.Foreground = Brushes.Red;
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
