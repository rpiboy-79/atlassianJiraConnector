# Atlassian Jira Power Query Connector

A custom Power Query (M) connector for Microsoft Power BI and Excel that integrates with Atlassian Jira Cloud. This README was updated to reflect recent improvements to the project's automated test harness, developer tooling, and developer-friendly helper scripts.

Version: 1.0.9 (connector Version updated in code)

## TL;DR — What's new

- Gateway and PowerBI Service Support
  - in liue of only the company identifier, you must now supply the full https:// string;
    - ```https://<companyIdentifier>.atlassian.net```
  - infrastructure is built for Service; in testing the Service would not return the proper URI during the grant process when run from a Gateway resulting in a failure.
- OAuth Support
  - Retrieval of Insights from Jira Product Discovery projects
    - Only attempts Insight retrieval when OAuth is used for authentication and there is at least one JPD project found in the target proejct(s).
    - Insights load as a seperate table since it is a 1:* relationship from Issues to Insights.
    - Detailed Snippet Data (Insights) is stored as ADF, requires significant transformation to make ussable.
    - Insight code modularized into its own PQM.
  - Authentication workflow moved to its own PQM.
  - Proper handling of URL based on credential Kind.
  - You must define an Application in the Atlassian Developer Console.
    - The redirect uses the Microsoft provided URI: ```https://oauth.powerbi.com/views/oauthredirect.html```
    - When creating your own Atlassian Application, use the URI for Authorization.
  - The Client ID and Client Secret are stored as plain text files locally on the machine using the connector.
    - Name the files:
      - ```client_id.txt```
      - ```client_secret.txt```
    - Update without build process:
      - Browse to the *.mez connector file.
      - Open in Winzip (or other archive tool); or modify the file extension to *.zip, (revert for implementation/use in Power BI / Power BI Gateway)
      - Copy ID and Secret file into the zip (*.mez) archive
- Support for Custom Sort Order
  - Sort by any valid fields, ASC or DESC
  - Combine multiple sorts into a single statement.
- Support for custom JQL strings.
  - the JQL query will be wrapped in parenthesis and appended with an ```AND``` to the base query which includes the Jira Project key
  - **DO NOT** include a project arguement in the JQL this will either result in a failure or strange results.
  - make sure to quote ```"<Textstring>"``` text values.
  - when using an OR operator over different Fields each field arguement must be wrapped in parenthesis, e.g. ```(status="In Progress") OR (priority="Major")```
- Full test-harness improvements: `tests/RunPQSDKTestSuites.ps1` rewritten for PowerShell 5.1 compatibility, automatic SDK discovery, robust encoding, HTML report generation, optional MEZ deployment, and improved error handling.
- Improved API handling and test support inside the connector: new `JiraTest` (unpublished) entrypoint to support test overrides and mock-data injection.
- Several fixes and refactors in `jiraAPI.pqm` and `atlassianJiraConnector.pq` to improve error handling, and mock-data support.

## 🚀 Features

- **Direct Jira Integration**: Connect to any Jira Cloud instance using API tokens
- **Navigation Table**: Browse projects and issues through an intuitive folder structure
- **Field Customization**: Specify which Jira fields to retrieve for optimal performance
- **Project Filtering**: Filter data by specific Jira projects during connection
- **Custom JQL**:include custom JQL filters and sorting
- **Dynamic Field Expansion**: Automatically expands nested Jira field structures
- **JPD Insights Retrieval**: return Insights on Issues from JPD Projects
- **Phase 2 MVP**: Support for OAuth
- **Built-in support for mock/test data**: used by the test harness

## 📋 Prerequisites

- **Power BI Desktop** or **Excel with Power Query**
- **Jira Cloud Instance** with API access
- **Jira API Token** (generated from your Atlassian account)
- **Visual Studio Code** with Power Query SDK extension (for development)

## ⚙️ Installation (End Users)

1. Download the latest `.mez` file from the releases
2. Copy the `.mez` file to your Power BI custom connectors folder:
   - **Power BI Desktop**: `%USERPROFILE%\Documents\Power BI Desktop\Custom Connectors\`
3. Enable custom connectors in Power BI:
   - File → Options and Settings → Options → Security
   - Under "Data Extensions", select "Allow any extension to load without validation"
4. Restart Power BI Desktop

To help with step 2 during development, use the helper script:

```powershell
# Copy the built MEZ to Power BI Desktop connectors folder (example)
.\tests\Copy-MezFile.ps1 -SourceMezPath ".\bin\AnyCPU\Debug\atlassianJiraConnector.mez" -Verbose
```

## 🧑‍� Developer Setup

1. Clone this repository
2. Install [Power Query SDK for Visual Studio Code](https://marketplace.visualstudio.com/items?itemName=PowerQuery.vscode-powerquery-sdk)
3. Open the project folder in VS Code
4. Build the connector using `Ctrl+Shift+P` → "Power Query: Build"

## 🔧 Configuration

### Obtaining Jira API Credentials

1. **Log in to Atlassian**: Go to [id.atlassian.com](https://id.atlassian.com)
2. **Create API Token**: 
   - Click "Security" → "Create and manage API tokens"
   - Click "Create API token"
   - Give it a descriptive name (e.g., "Power BI Connector")
   - Copy the token (save it securely!)
3. **Note Your Details**:
   - **Email**: Your Atlassian account email
   - **Instance**: Your Jira URL (e.g., `yourcompany.atlassian.net`)
   - **API Token**: The token you just created

### Set-up credentials for OAuth ###

1. **log in to**: [https://developer.atlassian.com/](https://developer.atlassian.com/)
2. **Click on Create**: to create a new OAuth 2.0 integration
3. **Follow steps to**: name your App, set permissions and authorize
4. **This app requires the following permissions/(Classic) scopes**:
   - View Jira issue data
   - View user profiles
   - Create and manage issues
5. **Store ID & Key**: store ID and Key in a secure location for use by connector.

### Connection Parameters

When connecting through Power BI:

| Parameter | Description | Example |
|-----------|-------------|---------|
| **Company URL Identifier** | Your Jira instance name (without `.atlassian.net`) | `acmesoftware` |
| **Field List** (Optional) | Comma-separated list of Jira fields to retrieve | `id,key,summary,created,assignee` |
| **Project Keys** (Optional) | Comma-separated list of project keys to filter | `PROJECT1,PROJECT2` |
| **Max Number of Results** (Optional) | Numerical value for the number of results to return from each Jira project, default is 1,000 | `2000` |
| **JQL Query String** (optional) | Custom JQL query string, appended to base query with an ```AND``` wrapped in Parenthesis ```()```, **DO NOT** include a project arguement. | `status = "In Progress"` |
| **JQL Order By String** (optional) | Custom JQL ORDER BY string, automatically appended to end of JQL query. Supports ASC and DESC, multiple fields | '"created DESC, priority DESC"' |

### Authentication

- **Username**: Your Atlassian account email
- **Password**: Your API token (not your Atlassian password)

### 🧪 Testing — full details (new test harness)

This project now includes a complete, local test harness and utilities in the `tests/` folder to run the pre-built PQ/PQOut tests, generate readable HTML reports, and optionally deploy the `.mez` to the local Power BI Connector directory.

Files of interest
- `tests/RunPQSDKTestSuites.ps1` — Main test runner (rewritten for PS 5.1 compatibility, auto-discovery of PQTest.exe, improved parsing of PQOut results, HTML report generation, optional MEZ copy, and safer JSON handling). Supports `-CopyMez`, `-PrettyJson`, `-MaxParallel` and other flags.
- `tests/Copy-MezFile.ps1` — Locates built `.mez` and copies it to the Power BI custom connectors directory.
- `tests/TestSuites/../*.query.pq` — Tests executed by the test harness.
- `tests/MockData/*.json` — Example test payloads used by parameter queries for deterministic tests (e.g., `jira_api_response_acme_software_pb.json`, `jira_api_response_null.json`).
- `tests/ConnectorConfigs/atlassianJiraConnector/ParameterQueries/atlassianJiraConnector_MockData.parameterquery.pq` — Example parameter query that calls the unpublished `JiraTest` entrypoint to exercise connector logic with mock data.

Running tests locally

Basic (auto-discover SDK tools and run tests):

```powershell
# Run the test harness; if PQTest.exe can be discovered automatically it will be used
.\tests\RunPQSDKTestSuites.ps1 -TestSettingsDirectoryPath ".\tests\ConnectorConfigs\atlassianJiraConnector\Settings" -TestSettingsList "SanitySettings.json" -PrettyJson -CopyMez
```

If you want to pass an explicit PQTest.exe path or the MEZ path, do so with the named parameters:

```powershell
.\tests\RunPQSDKTestSuites.ps1 -PQTestExePath 'C:\path\to\PQTest.exe' -ExtensionPath '.\bin\AnyCPU\Debug\atlassianJiraConnector.mez' -TestSettingsDirectoryPath '.\tests\ConnectorConfigs\atlassianJiraConnector\Settings' -TestSettingsList 'SanitySettings.json' -CopyMez
```

What the harness does
- Locates PQTest.exe automatically (from VS Code Power Query SDK extension) when possible.
- Runs the PQ/PQOut tests (the repository contains many pre-built `.query.pq` and `.query.pqout` files under `tests/TestSuites`).
- Parses `.pqout` results into HTML-friendly tables. The new parser handles simple string results and table `#table(...)` structures robustly.
- Produces an HTML report (and keeps historical copies) suitable for CI artifacts or local review.
- Optionally copies the built `.mez` into the user's Power BI Custom Connectors directory so tests can execute against the actual connector binary.
- Supports mock-data injection through the `JiraTest` entrypoint and the `TestDataOverride` mechanism so tests are deterministic and offline-capable.

Test configuration
- Configure the test harness by editing `tests/RunPQSDKTestSuitesSettings.json` or passing parameters to the runner.
- The connector parameter query in `tests/ConnectorConfigs/.../atlassianJiraConnector_MockData.parameterquery.pq` demonstrates how to hand a JSON object into the connector using `JiraTest(...)` for SDK-driven tests.

Troubleshooting the test harness
- If PQTest.exe cannot be found automatically, pass `-PQTestExePath` with the explicit path to the SDK tool.
- If encoding issues appear in PS 5.1, the runner uses a UTF8-no-BOM write mechanism to ensure the HTML and JSON artifacts are portable.

CI Suggestions

- Add a CI job that runs the test harness script on a Windows runner. Example pipeline steps:
  - Build (Power Query: Build) → produce `.mez`
  - Run tests: execute `RunPQSDKTestSuites.ps1` with `-ExtensionPath` pointing to the built `.mez` and collect the HTML output as artifacts
  - Optionally fail the job on non-empty failures by parsing the produced JSON/HTML summary

#### Test-driven developer workflow (recommended)

1. Implement a function change in `atlassianJiraConnector.pq` or supporting `.pqm` modules.
2. Add or update a `.query.pq` test under `tests/TestSuites` that exercises the change.
3. Use `tests/ConnectorConfigs/.../ParameterQueries` to create a deterministic test scenario (use `JiraTest` + mock JSON files when possible).
4. Run `RunPQSDKTestSuites.ps1` locally and verify the HTML report.
5. Iterate until tests pass.

## Connector internals (brief)

- Main logic: `atlassianJiraConnector.pq` (navigation table, `jiraQuery`, `jiraNavTbl`).
- API helpers: `jiraAPI.pqm` contains core retrieval helpers and parsing improvements. The code base now better handles nested/double-nested issue arrays and performs safer JSON extraction.
- Test entrypoint: `JiraTest(...)` — an unpublished, test-only entry function that accepts `TestDataOverride` and other parameters to exercise connector code paths with injected mock data.

## 📊 Usage Examples

### Basic Connection
Company URL Identifier: mycompany
Field List: (leave blank for default fields) - id, key, summary, description, issuetype, created, components
Project Keys: (leave blank for all projects)
Max Results: (leave blank to return 1,000 rows)
JQL Query String: (leave blank to return the top N rows from each project)

### Specific Projects with Custom Fields
Company URL Identifier: mycompany
Field List: id,key,summary,status,assignee,created,updated
Project Keys: PROJ1,PROJ2,WEBDEV
JQL Query String: (status="In Progress") OR (priority="Major")
JQL Order By String: created DESC, priority DESC

## 🏗️ Architecture

### Project Structure

'''
atlassianJiraConnector/
├── README.md # this document
├── atlassianJiraConnector.pq # Main connector logic
├── atlassianJiraConnector.proj # Project file
├── atlassianJiraConnector.query.pq # Test queries (legacy)
├── resources.resx # Localization resources
├── *.png # Connector icons
├── bin/
│ |── AnyCPU/
│ |── Debug/
│ └── atlassianJiraConnector.mez # Compiled connector
|── released/
│ |── version1.0/
│ └── atlassianJiraConnector.mez # Officially published connector
|── testResultServer/
│ └── dashboard_server.py # Mini python webservice/server for unit testing dashboard
└── tests/
  |── ConnectorConfigs/ # Configuration files for various unit tests
  |── Diagnostics/
  |── MockData/ # Mock data & configuration files for unit testing
  |── TestSuites/ # Various Unit Test Suites
  |── RunPQSDKTestSuites.ps1 # Cusomtimzed script for running unit tests and generating HTML results for use with webservice
  └── local-dashboard.html # Reporting dashboard for unit tests, requires dashboard_server.py to be running
'''

### Key Components

#### Data Retrieval
- **Phase 1**: `jiraDataRetrievalSimple()` - Single API calls (current)
- **Phase 2**: `jiraDataRetrievalPaginated()` - Full pagination (OAuth)

#### Core Functions
- **`jiraQuery()`**: Universal query function with transformation
- **`jiraNavTbl()`**: Navigation table generator
- **`createIssuesNavigationTable()`**: Sub-navigation builder

## 📈 Performance Considerations

### Current Limitations (Phase 1)
- **Maximum ~1000-2000 issues per project** (API token limitation)
- **No automatic pagination** (single large requests)
- **Sequential project processing** (not parallelized)

### Optimization Tips
- Use **Project Keys** parameter to limit scope
- Specify **Field List** to reduce payload size
- Consider breaking large projects into smaller queries

### Phase 2 Improvements (Planned)
- **OAuth 3LO Authentication** for unlimited pagination
- **Jira Product Discovery Insights** retrieve via Polaris Graphql endpoints
- **Incremental refresh support**

## 🔒 Security

### Credential Management
- **Never commit credentials** to version control
- Use **secure credential storage** for testing
- **API tokens** are preferred over passwords
- Consider **token rotation policies**

### Best Practices
- **Limit API token scope** to necessary permissions
- **Monitor API usage** to avoid rate limits
- **Use project filtering** to minimize data exposure
- **Regular security audits** of access patterns

## 🐛 Troubleshooting

### Common Issues

#### "Cannot connect to data source"
- Verify your Jira instance URL
- Check API token validity
- Ensure your account has project access

#### "No data returned"
- Verify project keys exist and are accessible
- Check field names are correct
- Try without field filtering first

#### "Pagination not working"
- This is expected in Phase 1 with API tokens
- Reduce `maxResults` or use project filtering
- Phase 2 (OAuth) will resolve this

### API Rate Limits

Jira Cloud API limits:
- **10,000 requests per hour** per API token
- **300 requests per minute** burst limit

## 🤝 Contributing

### Development Setup

1. **Fork and clone** the repository
2. **Install Power Query SDK** extension in VS Code
3. **Create feature branch**: `git checkout -b feature/amazing-feature`
4. **Add tests** for new functionality
5. **Ensure all tests pass**
6. **Submit pull request**

### Code Standards

- **Follow M language conventions**
- **Add comprehensive error handling**
- **Include unit tests** for new functions
- **Update documentation** for API changes
- **Maintain backward compatibility**

### Testing New Features

Build connector
Power Query: Build

Test in Power BI or write *.query.pq tests to run locally with command "Evaulate Current Power Query file."

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙋‍♂️ Support

### Getting Help

- **GitHub Issues**: For bugs and feature requests
- **Discussions**: For questions and community support
- **Wiki**: For detailed documentation and examples

### Reporting Issues

When reporting issues, please include:
- Power BI Desktop version
- Jira Cloud instance details (no credentials!)
- Error messages (full text)
- Steps to reproduce
- Expected vs actual behavior

## 🗺️ Roadmap

### Phase 1 ✅ (Complete)
- [x] Basic API token authentication
- [x] Project and issue navigation
- [x] Field customization
- [x] Comprehensive test suite
- [x] Documentation

### Phase 2 🚧 (Planned/Active)
- [x] OAuth 3LO authentication implementation* (unsecure)
- [ ] Full pagination support
- [x] Jira Product Discovery (JPD) Insights retreival
- [ ] Performance optimizations
- [x] Advanced filtering options

### Phase 3 🔮 (Future)
- [ ] Incremental refresh capabilities
- [ ] Advanced JQL query builder
- [ ] Examples & improved documentation
---

## 📞 Contact

**Project Maintainer**: Robert Manna  
**Organization**: Drofus AS  
**Repository**: [GitHub Repository URL]

---

*Last updated: November 10, 2025*
