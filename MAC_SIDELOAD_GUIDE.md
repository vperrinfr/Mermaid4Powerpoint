# Sideloading Guide for PowerPoint on Mac

If you're having trouble selecting the manifest.xml file in PowerPoint for Mac, use this alternative method.

## Method 1: Using the wef Folder (Recommended for Mac)

PowerPoint for Mac requires manifests to be placed in a specific folder.

### Step 1: Create the wef Folder

Open Terminal and run:

```bash
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
```

### Step 2: Copy the Manifest

Copy your manifest.xml to the wef folder:

```bash
cp manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
```

### Step 3: Restart PowerPoint

1. Quit PowerPoint completely (Cmd+Q)
2. Reopen PowerPoint
3. Create or open a presentation

### Step 4: Verify the Add-in Loaded

1. Go to **Insert** tab
2. Click **My Add-ins** dropdown
3. Look for **"Mermaid Diagrams"** under **Developer Add-ins**
4. Click on it to open the task pane

## Method 2: Using a Script (Automated)

I'll create a script to automate this process.

### Step 1: Make the Script Executable

```bash
chmod +x sideload-mac.sh
```

### Step 2: Run the Script

```bash
./sideload-mac.sh
```

The script will:
- Create the wef folder if it doesn't exist
- Copy the manifest
- Tell you to restart PowerPoint

## Troubleshooting

### Issue: "Developer Add-ins" section is empty

**Solution:**
1. Ensure the dev server is running (`npm start`)
2. Check that manifest.xml was copied correctly:
   ```bash
   ls -la ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
   ```
3. Verify the manifest.xml contains the correct localhost URL (https://localhost:3000)

### Issue: Certificate/SSL errors

**Solution:**
1. Ensure SSL certificates are installed:
   ```bash
   npx office-addin-dev-certs install
   ```
2. Trust the certificate in Keychain Access:
   - Open **Keychain Access** app
   - Search for "localhost"
   - Double-click the certificate
   - Expand **Trust** section
   - Set "When using this certificate" to **Always Trust**
   - Close and enter your password

### Issue: Add-in appears but won't load

**Solution:**
1. Check that the dev server is running on port 3000
2. Open Safari and navigate to https://localhost:3000/taskpane.html
3. Accept the certificate warning
4. Return to PowerPoint and try again

### Issue: "Cannot connect to localhost"

**Solution:**
1. Verify the server is running:
   ```bash
   lsof -i :3000
   ```
2. If nothing is running, start it:
   ```bash
   npm start
   ```
3. Wait for "Compiled successfully" message

## Alternative: Using PowerPoint Online

If Mac sideloading continues to have issues, you can use PowerPoint Online:

1. Go to https://office.com
2. Sign in with your Microsoft account
3. Open PowerPoint Online
4. Create or open a presentation
5. Go to **Insert** → **Add-ins** → **More Add-ins**
6. Click **Upload My Add-in**
7. Select your manifest.xml file
8. Click **Upload**

This method is more reliable and doesn't require the wef folder.

## Verification Steps

After sideloading, verify everything works:

1. ✅ Add-in appears in Insert → My Add-ins
2. ✅ Task pane opens when clicked
3. ✅ Interface loads without errors
4. ✅ Example diagrams render in preview
5. ✅ Can insert diagram into slide

## Quick Reference Commands

```bash
# Create wef folder
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef

# Copy manifest
cp manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/

# Check if manifest is there
ls -la ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/

# Remove manifest (if needed)
rm ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml

# Check if server is running
lsof -i :3000

# Start server
npm start

# Install certificates
npx office-addin-dev-certs install
```

## Need More Help?

If you're still having issues:

1. Check the browser console in PowerPoint (if available)
2. Look at the terminal output where `npm start` is running
3. Try PowerPoint Online as an alternative
4. Ensure your PowerPoint version supports add-ins (2016 or later)

## Mac-Specific Notes

- PowerPoint for Mac may cache add-ins. If changes don't appear, clear the cache:
  ```bash
  rm -rf ~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/*
  ```
- Some Mac versions require additional security permissions for localhost
- The wef folder method is more reliable than the UI upload on Mac
- You may need to approve the add-in in System Preferences → Security & Privacy

---

**Still having issues?** Open an issue on GitHub with:
- Your macOS version
- PowerPoint version
- Error messages
- Terminal output