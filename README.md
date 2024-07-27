# Excel ChatGPT Integration

This repository contains a VBA script that enables Excel to interact with the OpenAI ChatGPT API. The script implements a custom Excel function, `GPT()`, which allows you to send queries to the ChatGPT model and receive responses directly within your Excel sheets.

## Prerequisites

Before you can use the `GPT()` function in Excel, ensure the following requirements are met:

- Microsoft Excel 2016 or later with VBA support.
- An active OpenAI API key.

## Setup Instructions

Follow these steps to set up and use the ChatGPT Excel integration:

### Step 1: Enable Developer Mode in Excel

To access the VBA tools needed for this project, you must first enable Developer Mode:

1. Open Excel.
2. Go to `File` > `Options` > `Customize Ribbon`.
3. In the right column, ensure the "Developer" checkbox is selected.
4. Click `OK` to save your settings.

### Step 2: Import the VBA Script

1. Open Excel and access the VBA editor by pressing `ALT + F11`.
2. Right-click on `VBAProject (YourWorkbookName)` in the left sidebar, select `Insert`, and then `Module`.
3. Import the `ChatGPT.bas` file into the module.

### Step 3: Add Required References

1. In the VBA editor, go to `Tools` > `References`.
2. Check `Microsoft Scripting Runtime` to enable dictionary and collection support.
3. Ensure that `Microsoft XML, v6.0` (or similar version) is checked to handle HTTP requests.

### Step 4: Install and Reference VBA-JSON

The script requires VBA-JSON for parsing JSON responses from the ChatGPT API. Follow these instructions to include it:

1. Download the latest version of VBA-JSON from [VBA-JSON on GitHub](https://github.com/VBA-tools/VBA-JSON/releases).
2. Import the `JsonConverter.bas` file into your Excel VBA project via the VBA editor.
3. In the VBA editor, go to `Tools` > `References` and add a reference to `Microsoft Scripting Runtime` if not already added, as VBA-JSON requires this.

### Step 5: Secure Your API Key

Store your OpenAI API key securely and include it in the VBA script:
```vba
Dim apiKey As String
apiKey = "Your_OpenAI_API_Key"  // Replace with your actual OpenAI API key
```

### Step 6: Test the Function

1. Enter a test query in a cell in Excel, for example:

```excel
=GPT("What is the capital of France?")
```

2. Ensure macros are enabled when you open the workbook.

### Step 7: Error Handling

The script includes basic error handling to manage potential issues with network requests or API responses. If an error occurs, the function will return a descriptive error message.

## Using the `GPT()` Function

You can use the GPT() function anywhere in your Excel workbook:

```excel
=GPT("Your question here")
=GPT("What is the capital city of " & A1)
```

This function concatenates text with the contents of cell A1 to form the query.

## Contributing

Contributions are welcome. Please fork the repository and submit pull requests with your improvements.

## License

MIT License

## Acknowledgments

- OpenAI for the API.

- Contributors to the VBA-JSON project for JSON parsing capabilities.

