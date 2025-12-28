using System;
using TimesheetGenerator.Maui;

namespace TimesheetGenerator.Maui;

public partial class MainPage : ContentPage
{
    private readonly TimesheetService _service;

    public MainPage()
    {
        InitializeComponent();
        _service = new TimesheetService();
        PckMonth.SelectedIndex = 0; // Default to January
    }

    private async void OnGenerateClicked(object sender, EventArgs e)
    {
        try
        {
            LblStatus.Text = "Generating...";
            LblStatus.TextColor = Colors.Orange;

            // 1. Validate Year
            if (!int.TryParse(EntYear.Text, out int year) || year < 2000 || year > 2100)
            {
                await DisplayAlert("Error", "Please enter a valid year (2000-2100).", "OK");
                LblStatus.Text = "Error";
                LblStatus.TextColor = Colors.Red;
                return;
            }

            // 2. Validate Month
            if (PckMonth.SelectedIndex < 0)
            {
                await DisplayAlert("Error", "Please select a month.", "OK");
                LblStatus.Text = "Error";
                LblStatus.TextColor = Colors.Red;
                return;
            }
            int month = PckMonth.SelectedIndex + 1;

            // 3. Get Approver
            string approverName = "";
            string approverRole = "";

            if (RbApprover1.IsChecked)
            {
                approverName = "Salem Jaber Al-Marri";
                approverRole = "Head of E-Portal Development";
            }
            else if (RbApprover2.IsChecked)
            {
                approverName = "Shaikha Abdulrahman AlKhulaifi";
                approverRole = "Digital Government Systems Expert";
            }

            // 4. Generate
            string filePath = await _service.GenerateTimesheetAsync(year, month, approverName, approverRole);

            LblStatus.Text = "Ready";
            LblStatus.TextColor = Colors.Gray;

            await DisplayAlert("Success", $"Timesheet generated successfully!\nPath: {filePath}", "OK");
        }
        catch (Exception ex)
        {
            LblStatus.Text = "Error";
            LblStatus.TextColor = Colors.Red;
            await DisplayAlert("Error", $"An error occurred: {ex.Message}", "OK");
        }
    }
}
