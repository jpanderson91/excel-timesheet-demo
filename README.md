# Excel Timesheet Demo: Automating Resource Tracking and Reporting

## Project Story

This project began with a common challenge: managing and reporting on timesheet data for 100–150 contract resources, each submitting their own Excel file every month. The manual process was time-consuming, error-prone, and made it difficult to get real-time insights into funding, headcount, and resource status.

### Step 1: Simulating the Existing Process
We created sample timesheet files for three resources (A, B, and C), each representing a typical monthly submission. These files include columns for resource name, month, hours worked, projected hours, allocated funding, actual spend, status, and type (hiring, promotion, replacement).

### Step 2: Consolidation
To demonstrate the improved process, we consolidated all individual timesheets into a single CSV file (`Consolidated_TimesheetData.csv`). This file serves as the unified data source for reporting and analysis.

### Step 3: Dashboard Automation
We developed a set of Excel VBA macros:
- **CreateDashboardMacro**: Automatically builds a dashboard with Pivot Tables and Charts for total/available funding, breakdowns by type, headcount by status, and projections vs. actuals.
- **EnhanceDashboardMacro**: Adds slicers for interactive filtering, calculated fields (like available funding), and descriptive chart titles.

### Step 4: Version Control and Collaboration
The entire project folder, including all sample data and macros, was initialized as a Git repository and pushed to GitHub. This enables easy sharing, versioning, and collaboration.

## Demo Walkthrough
1. **Show the sample timesheet files** to illustrate the original manual process.
2. **Open the consolidated data file** to demonstrate how all timesheets are unified for analysis.
3. **Run the dashboard macros in Excel** to instantly generate interactive reports and visualizations.
4. **Use slicers and charts** to filter and explore the data by resource, status, and month.
5. **Highlight the automation and repeatability**: New timesheets can be added, and the dashboard refreshed with a single click.
6. **Show the GitHub repo** to explain how the process is now documented, versioned, and ready for team collaboration.

## Files in This Repo
- `Timesheet_ResourceA_July2025.csv`, `Timesheet_ResourceB_July2025.csv`, `Timesheet_ResourceC_July2025.csv`: Sample individual timesheets
- `Consolidated_TimesheetData.csv`: Unified data for reporting
- `CreateDashboardMacro.bas`, `EnhanceDashboardMacro.bas`: VBA macros for dashboard automation
- `dashboard example.png`: Screenshot of the generated dashboard

## Next Steps
- Add more sample data to simulate a full team
- Customize the dashboard for your organization’s needs
- Use GitHub for ongoing improvements and collaboration

---

This project demonstrates how manual, spreadsheet-based processes can be transformed into automated, interactive, and collaborative solutions using Excel and GitHub.
