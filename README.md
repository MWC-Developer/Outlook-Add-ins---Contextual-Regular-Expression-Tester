# Outlook Contextual Regular Expression Tester

A contextual Outlook add-in that demonstrates regular expression matching against email message bodies. This add-in provides a task pane interface for retrieving and testing regular expressions against message content in both plain text and HTML formats.

## Features

### Contextual Activation
The add-in activates automatically when reading messages that contain specific patterns matching predefined regular expressions:
- **Rule1**: `\bS[2-3]\d{9}\b` - Matches patterns like S2000000001, S3999999999
- **Rule2**: `\bB[A-Z]{3}\d{6}\b` - Matches patterns like BABC123456
- **Rule3**: Complex pattern matching various formats including alphanumeric codes

### Task Pane Functionality
- **Message Body Retrieval**: Retrieve the current message body in Plain Text or HTML format
- **RegEx Testing**: Test custom regular expressions against the retrieved message body
- **Pinnable Task Pane Support**: Task pane remains open when switching between messages with automatic content updates
- **Real-time Validation**: Test RegEx button enables/disables based on input validity

## Setup and Development

### Prerequisites
- Outlook (Desktop, Web, or Mac)
- A web server to host the add-in files (IIS, Node.js, Python HTTP server, etc.)
- For sideloading: Access to Outlook add-in management

### Configuration

#### 1. Update the Manifest
The manifest file `RegEx Sample Add-in (generic).xml` uses `~remoteUrl` as a placeholder for your hosting location. Replace all instances with your actual server URL:

```xml
<!-- Replace ~remoteUrl with your server URL, e.g., https://localhost:3000 or https://yourdomain.com -->
```

**Find and replace:**
- `~remoteUrl` → `https://your-server-url`

Example locations in the manifest:
- `<IconUrl>`: Image resources
- `<SourceLocation>`: Task pane HTML file
- `<AppDomain>`: Your domain for security
- `<bt:Image>`: Icon resources
- `<bt:Url>`: All HTML page references

#### 2. Host the Add-in Files

**Option A: Local Development (IIS)**
1. Place all files in an IIS virtual directory or application
2. Ensure HTTPS is configured (required for Outlook add-ins)
3. Update manifest with `https://localhost` or your local IIS URL

**Option B: Local Development (Node.js)**
```bash
# Install http-server globally
npm install -g http-server

# Navigate to the RegExSample folder
cd c:\Apps\WebAPISample\wwwroot\RegExSample

# Start server with CORS enabled
http-server -p 3000 --cors -c-1
```
Update manifest with `https://localhost:3000`

**Option C: Production Server**
1. Upload all files to your web server
2. Ensure HTTPS is enabled (required)
3. Update manifest with your production URL
4. Configure CORS headers if needed

#### 3. Sideload the Add-in

**Outlook Desktop (Windows):**
1. Save the updated manifest XML file
2. Open Outlook
3. Go to File → Manage Add-ins (or Get Add-ins) - this will open OWA to the add-in installation screen
4. Click "My add-ins" → "Add a custom add-in" → "Add from file"
5. Browse to your manifest file and install

**Outlook Web:**
1. Click Settings (gear icon) → View all Outlook settings
2. Navigate to Mail → Customize actions → Custom add-ins
3. Click "Add a custom add-in" → "Add from file"
4. Upload your manifest file

**For detailed instructions:** [Sideload Outlook add-ins for testing](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing)

### Testing the Add-in

1. **Contextual Activation**: Send yourself an email containing one of the predefined patterns (e.g., "S2123456789")
2. Open the message in Outlook
3. The add-in should activate automatically and appear in the ribbon
4. Click the add-in button to open the task pane

**Task Pane Features:**
- Select Plain Text or HTML format
- Click "Retrieve message body" to load the content
- Enter a regular expression pattern in the input field
- Click "Test RegEx" to find matches (button enables when valid regex is entered)
- Results display above the message body showing all matches

### File Structure

```
RegExSample/
├── RegEx Sample Add-in (generic).xml    # Manifest file (update ~remoteUrl)
├── MessageRead.html                      # Task pane HTML
├── MessageRead.js                        # Task pane JavaScript logic
├── MessageRead.css                       # Task pane styles
├── index.html                            # Detected entity page
├── index.js                              # Detected entity logic
├── index.css                             # Detected entity styles
├── Functions/
│   ├── FunctionFile.html                # Function file (for commands, empty as not implemented)
│   └── FunctionFile.js                  # Function file logic
└── images/                               # Add-in icons (various sizes)
```

## Key Features Implementation

### Pinnable Task Pane
The add-in supports pinnable task panes using the `ItemChanged` event:
```javascript
Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
```
When a message changes and the task pane was previously showing content, it automatically refreshes.

### RegEx Validation
The add-in validates regular expressions before enabling the test button:
```javascript
try {
  new RegExp(regexInput.value);
  hasValidRegex = true;
} catch (error) {
  hasValidRegex = false;
}
```

### Body Retrieval
Supports both plain text and HTML formats using Office.js coercion types:
```javascript
Office.context.mailbox.item.body.getAsync(
  Office.CoercionType.Text, // or Office.CoercionType.Html
  callback
);
```

## Technologies Used

- **Office.js**: Microsoft Office JavaScript API for Outlook integration
- **Fluent UI**: Microsoft's design system for consistent styling
- **Vanilla JavaScript**: No framework dependencies for task pane logic

## License

Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.

## Additional Resources

- [Outlook Add-ins Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/)
- [Pinnable Task Pane](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/pinnable-taskpane)
- [Activation Rules](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/activation-rules)
- [Regular Expression Activation](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/use-regular-expressions-to-show-an-outlook-add-in)
