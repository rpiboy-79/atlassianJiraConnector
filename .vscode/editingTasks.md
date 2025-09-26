# To extend the automated testing tasks:

[Thread](https://www.perplexity.ai/search/are-you-familiar-with-the-micr-tC0msPRcTIamksq4U.DkUQ?30=d#61)

## Easy to Extend:
Want to add performance tests later? Just add:

**json**
```
{
    "label": "Run Performance Tests (Automated)",
    "type": "shell",
    "command": "powershell.exe",
    "args": [
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "${workspaceFolder}/tests/RunPQSDKTestSuites.ps1",
        "-TestSettingsList", "@('PerformanceTests.json')",
        "-HTMLPreview",
        "-SkipConfirmation",
        "-NoOpenHTML"
    ],
    "group": "test",
    "options": {
        "cwd": "${workspaceFolder}/tests"
    },
    "problemMatcher": []
}
Then update the sequence:

json
"dependsOn": [
    "Run Basic Connection Tests (Automated)",
    "Run Mock Data Tests (Automated)",
    "Run Performance Tests (Automated)"
]
```
This approach is much more maintainable and provides better developer experience!