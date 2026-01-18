# Setup Guide for Mermaid PowerPoint Plugin

This guide will walk you through setting up and running the Mermaid PowerPoint Plugin.

## Prerequisites

Before you begin, ensure you have the following installed:

1. **Node.js** (v16 or higher)
   - Download from: https://nodejs.org/
   - Verify installation: `node --version`

2. **npm** (comes with Node.js)
   - Verify installation: `npm --version`

3. **PowerPoint**
   - Microsoft 365, PowerPoint 2016, or later
   - PowerPoint for Mac or Windows

## Step-by-Step Installation

### 1. Install Dependencies

Open a terminal in the project directory and run:

```bash
npm install
```

This will install all required packages including:
- React and TypeScript
- Mermaid.js for diagram rendering
- Fluent UI components
- Office.js for PowerPoint integration
- Webpack for bundling

### 2. Generate SSL Certificates

Office Add-ins require HTTPS for security. Generate local SSL certificates:

```bash
npx office-addin-dev-certs install
```

**Note for macOS users**: You may need to enter your password to trust the certificate.

**Note for Windows users**: You may see a security warning - click "Yes" to install the certificate.

### 3. Update Manifest (Optional)

If you want to customize the add-in:

1. Open `manifest.xml`
2. Update the `<Id>` with a unique GUID (you can generate one at https://guidgenerator.com/)
3. Update URLs if not using localhost:3000

### 4. Start Development Server

Start the webpack dev server:

```bash
npm start
```

You should see output indicating the server is running on `https://localhost:3000`

**Keep this terminal window open** - the server needs to run while you use the add-in.

### 5. Sideload the Add-in in PowerPoint

#### For PowerPoint on Windows:

1. Open PowerPoint
2. Create a new presentation or open an existing one
3. Go to **Insert** tab → **Get Add-ins** (or **My Add-ins**)
4. Click **Upload My Add-in** at the bottom
5. Browse to your project folder and select `manifest.xml`
6. Click **Upload**

#### For PowerPoint on Mac:

1. Open PowerPoint
2. Go to **Insert** tab → **Add-ins** → **My Add-ins**
3. Click the dropdown next to **My Add-ins** and select **Add from File...**
4. Browse to your project folder and select `manifest.xml`
5. Click **Add**

#### For PowerPoint Online:

1. Open PowerPoint Online (office.com)
2. Create or open a presentation
3. Go to **Insert** tab → **Add-ins** → **More Add-ins**
4. Click **Upload My Add-in**
5. Browse to your project folder and select `manifest.xml`
6. Click **Upload**

### 6. Use the Add-in

Once loaded, you should see:

1. A new **"Mermaid Diagrams"** group in the **Home** tab
2. A **"Create Diagram"** button in that group
3. Click the button to open the task pane

## Troubleshooting

### Issue: "Add-in won't load" or "Manifest error"

**Solution:**
1. Validate your manifest:
   ```bash
   npm run validate
   ```
2. Ensure the dev server is running (`npm start`)
3. Check that URLs in manifest.xml match your server (default: https://localhost:3000)

### Issue: "Certificate not trusted" error

**Solution:**
1. Reinstall certificates:
   ```bash
   npx office-addin-dev-certs install --force
   ```
2. Restart PowerPoint
3. On macOS, you may need to manually trust the certificate in Keychain Access

### Issue: "Cannot find module" errors

**Solution:**
1. Delete `node_modules` and `package-lock.json`:
   ```bash
   rm -rf node_modules package-lock.json
   ```
2. Reinstall dependencies:
   ```bash
   npm install
   ```

### Issue: Port 3000 already in use

**Solution:**
1. Change the port in `webpack.config.js`:
   ```javascript
   devServer: {
     port: 3001, // Change to any available port
     // ...
   }
   ```
2. Update all URLs in `manifest.xml` to use the new port
3. Restart the dev server

### Issue: Diagram doesn't render

**Solution:**
1. Check the browser console for errors (F12 in PowerPoint)
2. Verify Mermaid syntax is correct
3. Try one of the example diagrams from the dropdown

### Issue: Can't insert diagram into slide

**Solution:**
1. Ensure you have at least one slide in your presentation
2. Check that the diagram rendered successfully in the preview
3. Try creating a new slide and inserting there

## Development Workflow

### Making Changes

1. Edit files in the `src/` directory
2. Webpack will automatically rebuild (hot reload enabled)
3. Refresh the task pane in PowerPoint to see changes
   - Right-click in the task pane → **Reload**
   - Or close and reopen the task pane

### Building for Production

When ready to deploy:

```bash
npm run build
```

This creates optimized files in the `dist/` directory.

### Testing

1. Test with different diagram types (flowchart, sequence, class)
2. Test with complex diagrams
3. Test inserting multiple diagrams
4. Test on different PowerPoint versions if possible

## Next Steps

1. **Customize the UI**: Edit `src/taskpane/App.tsx` and `src/taskpane/App.css`
2. **Add More Diagram Types**: Extend the `defaultDiagrams` object in `App.tsx`
3. **Improve Styling**: Modify Mermaid theme in the `mermaid.initialize()` call
4. **Add Features**: Implement diagram editing, templates, or export options

## Useful Commands

```bash
# Start development server
npm start

# Build for production
npm run build

# Validate manifest
npm run validate

# Open in browser (for debugging)
npm run dev
```

## Additional Resources

- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Mermaid.js Documentation](https://mermaid.js.org/)
- [Fluent UI React Components](https://react.fluentui.dev/)
- [TypeScript Documentation](https://www.typescriptlang.org/docs/)

## Getting Help

If you encounter issues:

1. Check the browser console (F12) for error messages
2. Review the terminal output where `npm start` is running
3. Consult the README.md for additional information
4. Open an issue on GitHub with:
   - Your PowerPoint version
   - Operating system
   - Error messages
   - Steps to reproduce

## Security Notes

- The add-in runs in a sandboxed environment
- SSL certificates are only for local development
- For production, use proper SSL certificates from a trusted authority
- Never commit sensitive data or API keys to the repository

Happy diagramming! 🎨📊