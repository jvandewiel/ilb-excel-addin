# Simple Sharing Guide for ILB Excel Add-in

## Quick Deployment Steps

### 1. Prepare Files for Upload

After running `npm run build`, you need to upload these files from the `dist` folder to a web server:

**Required Files:**
```
dist/
├── taskpane.html
├── taskpane.js
├── commands.html  
├── commands.js
├── functions.html
├── functions.js
├── functions.json
├── polyfill.js
├── *.css files
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-64.png
    ├── icon-80.png
    ├── almelo.json
    └── renumerarions-export.json
```

### 2. Upload to Web Server

Choose one of these free hosting options:

#### Option A: GitHub Pages (Recommended)
1. Create a new GitHub repository
2. Upload all files from `dist` folder to the repository
3. Go to Settings > Pages
4. Select "Deploy from a branch" and choose "main"
5. Your URL will be: `https://USERNAME.github.io/REPOSITORY-NAME`

#### Option B: Netlify
1. Go to [netlify.com](https://netlify.com)
2. Drag and drop the `dist` folder
3. Get your URL (e.g., `https://amazing-name-12345.netlify.app`)

#### Option C: Vercel
1. Go to [vercel.com](https://vercel.com)
2. Import your project or upload files
3. Get your URL (e.g., `https://your-project.vercel.app`)

### 3. Update Manifest File

1. Copy `manifest-deployment-template.xml` to `manifest-production.xml`
2. Replace ALL instances of `https://YOUR-DOMAIN.com` with your actual URL
3. Example: If your URL is `https://myname.github.io/ilb-excel`, replace:
   ```xml
   https://YOUR-DOMAIN.com/taskpane.html
   ```
   with:
   ```xml
   https://myname.github.io/ilb-excel/taskpane.html
   ```

### 4. Share with Others

Send them:
1. The updated `manifest-production.xml` file
2. Installation instructions (see below)

## Installation Instructions for Recipients

### For Excel Desktop:
1. Open Excel
2. Go to **Insert** > **Get Add-ins** > **My Add-ins**
3. Click **Upload My Add-in**
4. Select the `manifest-production.xml` file
5. Click **Upload**

### For Excel Online:
1. Open Excel in browser
2. Go to **Insert** > **Add-ins** > **My Add-ins**
3. Click **Upload My Add-in**
4. Select the `manifest-production.xml` file
5. Click **Upload**

## Example Deployment URLs

If you deploy to GitHub Pages at `https://myname.github.io/ilb-excel/`, your manifest should have:

```xml
<IconUrl DefaultValue="https://myname.github.io/ilb-excel/assets/icon-32.png"/>
<SourceLocation DefaultValue="https://myname.github.io/ilb-excel/taskpane.html"/>
<bt:Url id="Taskpane.Url" DefaultValue="https://myname.github.io/ilb-excel/taskpane.html"/>
```

## Testing Your Deployment

1. Open your deployment URL + `/taskpane.html` in a browser
2. You should see the ILB tool interface
3. If it works in browser, it should work in Excel

## Troubleshooting

- **HTTPS Required**: All URLs must use HTTPS (not HTTP)
- **CORS**: Some hosting services handle this automatically
- **File Paths**: Ensure all file paths in manifest match your server structure
- **Icons**: Make sure icon files are accessible at the specified URLs

## File Structure on Server

Your web server should have this structure:
```
/
├── taskpane.html
├── taskpane.js  
├── commands.html
├── commands.js
├── functions.html
├── functions.js
├── functions.json
├── polyfill.js
├── [css files]
└── assets/
    ├── icon-16.png
    ├── icon-32.png  
    ├── icon-64.png
    ├── icon-80.png
    ├── almelo.json
    └── renumerarions-export.json
```

## Quick GitHub Pages Setup

1. Create new repo on GitHub
2. Upload `dist` folder contents
3. Enable Pages in repo settings
4. Update manifest with GitHub Pages URL
5. Share manifest file

Your add-in will be available at: `https://USERNAME.github.io/REPO-NAME/`