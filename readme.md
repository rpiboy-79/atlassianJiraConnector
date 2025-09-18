# Atlassian Jira Power Query Connector

A custom Power Query connector for Microsoft Power BI and Excel that enables direct integration with Atlassian Jira Cloud instances. This connector provides seamless access to Jira issues, projects, and related data for business intelligence and reporting purposes.

## 🚀 Features

- **Direct Jira Integration**: Connect to any Jira Cloud instance using API tokens
- **Navigation Table**: Browse projects and issues through an intuitive folder structure
- **Field Customization**: Specify which Jira fields to retrieve for optimal performance
- **Project Filtering**: Filter data by specific Jira projects during connection
- **Dynamic Field Expansion**: Automatically expands nested Jira field structures
- **Phase 1 MVP**: Works with API token authentication (Phase 2 will add OAuth 3LO support)

## 📋 Prerequisites

- **Power BI Desktop** or **Excel with Power Query**
- **Jira Cloud Instance** with API access
- **Jira API Token** (generated from your Atlassian account)
- **Visual Studio Code** with Power Query SDK extension (for development)

## ⚙️ Installation

### For End Users

1. Download the latest `.mez` file from the releases
2. Copy the `.mez` file to your Power BI custom connectors folder:
   - **Power BI Desktop**: `Documents\Power BI Desktop\Custom Connectors\`
3. Enable custom connectors in Power BI:
   - File → Options and Settings → Options → Security
   - Under "Data Extensions", select "Allow any extension to load without validation"
4. Restart Power BI Desktop

### For Developers

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

### Connection Parameters

When connecting through Power BI:

| Parameter | Description | Example |
|-----------|-------------|---------|
| **Company URL Identifier** | Your Jira instance name (without `.atlassian.net`) | `acmesoftware` |
| **Field List** (Optional) | Comma-separated list of Jira fields to retrieve | `id,key,summary,created,assignee` |
| **Project Keys** (Optional) | Comma-separated list of project keys to filter | `PROJECT1,PROJECT2` |
| **Max Number of Results** (Optional) | Numerical value for the number of results to return from each Jira project, default is 1,000 | '2000' |

### Authentication

- **Username**: Your Atlassian account email
- **Password**: Your API token (not your Atlassian password)

## 📊 Usage Examples

### Basic Connection
Company URL Identifier: mycompany
Field List: (leave blank for default fields) - id, key, summary, description, issuetype, created, components
Project Keys: (leave blank for all projects)

### Specific Projects with Custom Fields
Company URL Identifier: mycompany
Field List: id,key,summary,status,assignee,created,updated
Project Keys: PROJ1,PROJ2,WEBDEV

## 🏗️ Architecture

### Project Structure

'''atlassianJiraConnector/
├── README.md # this document
├── atlassianJiraConnector.pq # Main connector logic
├── atlassianJiraConnector.proj # Project file
├── atlassianJiraConnector.query.pq # Test queries (legacy)
├── resources.resx # Localization resources
├── *.png # Connector icons
├── bin/
│ └── AnyCPU/
│ └── Debug/
│ └── atlassianJiraConnector.mez # Compiled connector
└── released/
│ └──version1.0/
│ └──atlassianJiraConnector.mez # Officially published connector'''

### Key Components

#### Data Retrieval
- **Phase 1**: `jiraDataRetrievalSimple()` - Single API calls (current)
- **Phase 2**: `jiraDataRetrievalPaginated()` - Full pagination (OAuth)

#### Core Functions
- **`jiraQuery()`**: Universal query function with transformation
- **`jiraNavTbl()`**: Navigation table generator
- **`createIssuesNavigationTable()`**: Sub-navigation builder

## 🧪 Testing

Unit Tests are seperated as per Microsoft SDK.

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

### Phase 1 ✅ (Current)
- [x] Basic API token authentication
- [x] Project and issue navigation
- [x] Field customization
- [x] Comprehensive test suite
- [x] Documentation

### Phase 2 🚧 (Planned)
- [ ] OAuth 3LO authentication implementation
- [ ] Full pagination support
- [ ] Jira Product Discovery (JPD) Insights retreival
- [ ] Performance optimizations
- [ ] Advanced filtering options
- [ ] Examples & improved documentation

### Phase 3 🔮 (Future)
- [ ] Incremental refresh capabilities
- [ ] Advanced JQL query builder

---

## 📞 Contact

**Project Maintainer**: Robert Manna  
**Organization**: Drofus AS  
**Repository**: [GitHub Repository URL]

---

*Last updated: September 2025*
