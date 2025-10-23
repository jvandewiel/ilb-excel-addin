# GitHub Deployment Guide for ILB Excel Add-in

## 🌐 Deploy to GitHub Pages

### Prerequisites
- GitHub account
- Git repository for your project
- Built dist/ folder with your add-in files

### Step 1: Prepare Repository

1. **Initialize Git Repository** (if not already done):
```bash
git init
git add .
git commit -m "Initial commit: ILB Excel Add-in"
```

2. **Create GitHub Repository**:
   - Go to GitHub.com
   - Click "New Repository"
   - Name: `ilb-excel-addin`
   - Make it Public (required for GitHub Pages free tier)
   - Don't initialize with README (you already have files)

3. **Connect Local to Remote**:
```bash
git remote add origin https://github.com/jvandewiel/ilb-excel-addin.git
git branch -M main
git push -u origin main
```

### Step 2: Configure GitHub Pages

1. **Go to Repository Settings**:
   - Navigate to your repository on GitHub
   - Click "Settings" tab
   - Scroll to "Pages" section

2. **Configure Source**:
   - Source: "Deploy from a branch"
   - Branch: `main` or `gh-pages`
   - Folder: `/dist` (if using dist folder) or `/` (root)

3. **Build Configuration**:
   - GitHub will automatically build and deploy
   - Your add-in will be available at: `https://jvandewiel.github.io/ilb-excel-addin/`

### Step 3: Update Manifest for GitHub Pages

Create a GitHub-specific manifest file:

**File: `manifest-github.xml`**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" 
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" 
           xsi:type="TaskPaneApp">
  
  <Id>d8ef3416-b9be-410b-a03f-13ad1c27b3da</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>ILB Tools</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="ILB Excel Tool"/>
  <Description DefaultValue="Inlenersbeloning data loader for Excel with remuneration component analysis."/>
  
  <!-- GitHub Pages URLs -->
  <IconUrl DefaultValue="https://jvandewiel.github.io/ilb-excel-addin/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://jvandewiel.github.io/ilb-excel-addin/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://github.com/jvandewiel/ilb-excel-addin"/>
  
  <AppDomains>
    <AppDomain>https://jvandewiel.github.io</AppDomain>
  </AppDomains>
  
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://jvandewiel.github.io/ilb-excel-addin/taskpane.html"/>
  </DefaultSettings>
  
  <Permissions>ReadWriteDocument</Permissions>
  
  <!-- Rest of VersionOverrides section with GitHub URLs -->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://jvandewiel.github.io/ilb-excel-addin/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://jvandewiel.github.io/ilb-excel-addin/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://jvandewiel.github.io/ilb-excel-addin/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://github.com/jvandewiel/ilb-excel-addin"/>
        <bt:Url id="Commands.Url" DefaultValue="https://jvandewiel.github.io/ilb-excel-addin/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://jvandewiel.github.io/ilb-excel-addin/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Get started with ILB Excel Tool"/>
        <bt:String id="CommandsGroup.Label" DefaultValue="ILB Tools"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Show ILB Taskpane"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Load and analyze remuneration data with the ILB Excel Tool"/>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Click to open the ILB data loader"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

### Step 4: Automated GitHub Actions Deployment

Create `.github/workflows/deploy.yml`:

```yaml
name: Deploy to GitHub Pages

on:
  push:
    branches: [ main ]
  pull_request:
    branches: [ main ]

jobs:
  build-and-deploy:
    runs-on: ubuntu-latest
    
    steps:
    - name: Checkout
      uses: actions/checkout@v3
      
    - name: Setup Node.js
      uses: actions/setup-node@v3
      with:
        node-version: '18'
        cache: 'npm'
        
    - name: Install dependencies
      run: npm install
      
    - name: Build
      run: npm run build
      
    - name: Deploy to GitHub Pages
      uses: peaceiris/actions-gh-pages@v3
      if: github.ref == 'refs/heads/main'
      with:
        github_token: ${{ secrets.GITHUB_TOKEN }}
        publish_dir: ./dist
```

## 📤 Sharing Methods

### Method 1: Direct Manifest Link
Share the manifest URL directly:
```
https://jvandewiel.github.io/ilb-excel-addin/manifest-github.xml
```

Recipients can:
1. Open Excel Online or Desktop
2. Go to Insert > Add-ins > Upload My Add-in
3. Enter the manifest URL

### Method 2: Repository Clone
Share your repository URL:
```
https://github.com/jvandewiel/ilb-excel-addin
```

Recipients can:
1. Clone the repository
2. Run `npm install` and `npm run build`
3. Sideload the manifest locally

### Method 3: Release Packages
Create GitHub releases with packaged add-in:
1. Build your project: `npm run build`
2. Create zip file with dist/ contents
3. Create GitHub release with the zip file
4. Recipients download and extract

## 🔧 Build Scripts

Add these scripts to your `package.json`:

```json
{
  "scripts": {
    "build:github": "npm run build && node scripts/update-github-manifest.js",
    "deploy:github": "npm run build:github && gh-pages -d dist"
  }
}
```

## 🌐 Custom Domain (Optional)

If you have a custom domain:
1. Create `CNAME` file in repository root
2. Add your domain name to the file
3. Configure DNS CNAME record pointing to `jvandewiel.github.io`

## 📋 Checklist for GitHub Deployment

- [ ] Repository created and files pushed
- [ ] GitHub Pages configured in repository settings
- [ ] Manifest URLs updated to GitHub Pages URLs
- [ ] All assets (HTML, JS, CSS, icons) in dist/ folder
- [ ] HTTPS enabled (automatic with GitHub Pages)
- [ ] Manifest validated for GitHub deployment
- [ ] Test the add-in URL before sharing

## 🔒 Security Considerations

- GitHub Pages serves over HTTPS (required for Office Add-ins)
- Public repositories are visible to everyone
- Consider using private repository + GitHub Pro for sensitive data
- Remove any sensitive API keys or data from public repository

## 🆘 Troubleshooting

**Common Issues:**
- **404 errors**: Check file paths in manifest match actual file locations
- **CORS errors**: GitHub Pages automatically handles CORS for add-ins
- **Manifest validation**: Use Office Add-in Validator tool
- **Loading issues**: Check browser console for detailed error messages

## 📞 Support

For GitHub deployment issues:
- GitHub Pages documentation: https://docs.github.com/en/pages
- Office Add-ins documentation: https://docs.microsoft.com/en-us/office/dev/add-ins/