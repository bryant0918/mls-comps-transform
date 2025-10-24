# GitHub Pages Deployment Guide

This guide will help you deploy the Existing Comps Transformer to GitHub Pages so your brother can access it from anywhere.

## 📋 Prerequisites

- A GitHub account (free)
- Git installed on your computer

## 🚀 Step-by-Step Deployment

### 1. Create a GitHub Repository

1. Go to [GitHub](https://github.com) and log in
2. Click the **"+"** button in the top-right corner
3. Select **"New repository"**
4. Fill in the details:
   - **Repository name**: `existing-comps-transformer` (or any name you prefer)
   - **Description**: "Transform raw MLS data into formatted comps"
   - Choose **Public** (required for free GitHub Pages)
   - **DO NOT** initialize with README (we already have files)
5. Click **"Create repository"**

### 2. Push Your Code to GitHub

Open Terminal/Command Prompt in your project folder and run:

```bash
# Initialize git repository (if not already done)
git init

# Add all files
git add index.html styles.css transform.js .gitignore DEPLOYMENT.md

# Commit the files
git commit -m "Initial commit: Existing Comps Transformer web app"

# Add your GitHub repository as remote (replace YOUR_USERNAME and REPO_NAME)
git remote add origin https://github.com/YOUR_USERNAME/existing-comps-transformer.git

# Push to GitHub
git branch -M main
git push -u origin main
```

**Note**: Replace `YOUR_USERNAME` with your GitHub username and `existing-comps-transformer` with your repository name.

### 3. Enable GitHub Pages

1. Go to your repository on GitHub
2. Click **"Settings"** (top menu)
3. Scroll down and click **"Pages"** in the left sidebar
4. Under **"Source"**, select:
   - Branch: **main**
   - Folder: **/ (root)**
5. Click **"Save"**
6. Wait 1-2 minutes for deployment

### 4. Access Your Live Site

After deployment, GitHub will show you the URL:

```
https://YOUR_USERNAME.github.io/existing-comps-transformer/
```

Share this URL with your brother! 🎉

## 📱 Optional: Custom Domain

If you want a custom domain (like `comps.yourdomain.com`):

1. Buy a domain from a registrar (GoDaddy, Namecheap, Google Domains, etc.)
2. In your domain's DNS settings, add a CNAME record:
   - Name: `comps` (or `@` for root domain)
   - Value: `YOUR_USERNAME.github.io`
3. In GitHub Pages settings, add your custom domain
4. Enable "Enforce HTTPS"

More info: [GitHub Docs - Custom Domain](https://docs.github.com/en/pages/configuring-a-custom-domain-for-your-github-pages-site)

## 🔄 Updating the Site

Whenever you make changes to the code:

```bash
# Make your changes to index.html, styles.css, or transform.js

# Stage changes
git add .

# Commit changes
git commit -m "Description of your changes"

# Push to GitHub
git push
```

GitHub Pages will automatically redeploy (takes 1-2 minutes).

## 🧪 Testing Locally Before Deployment

To test the site on your computer before deploying:

### Option 1: Python Simple Server

```bash
# In the project folder
python3 -m http.server 8000
```

Then open: http://localhost:8000

### Option 2: VS Code Live Server

1. Install "Live Server" extension in VS Code
2. Right-click `index.html`
3. Select "Open with Live Server"

### Option 3: Node.js http-server

```bash
# Install globally (one time)
npm install -g http-server

# Run in project folder
http-server
```

## 🎨 Customization Ideas

### Change Colors
Edit `styles.css` - look for color values like `#667eea` and `#764ba2`

### Change Quartile Colors
Edit `transform.js` - look for the `QUARTILE_COLORS` object

### Change Project Name
Edit `index.html` and `transform.js` - search for "Pagoda Grove Circle"

### Add Logo
1. Add your logo file (e.g., `logo.png`) to the project
2. In `index.html`, add before the `<h1>` tag:
```html
<img src="logo.png" alt="Logo" style="width: 100px; margin-bottom: 20px;">
```

## 📊 Analytics (Optional)

To track usage, add Google Analytics:

1. Create a Google Analytics account
2. Get your tracking ID
3. Add this before `</head>` in `index.html`:

```html
<!-- Google Analytics -->
<script async src="https://www.googletagmanager.com/gtag/js?id=GA_MEASUREMENT_ID"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());
  gtag('config', 'GA_MEASUREMENT_ID');
</script>
```

## 🔒 Security & Privacy

The current implementation:
- ✅ All processing happens in the browser
- ✅ Files never leave the user's computer
- ✅ No data is sent to any server
- ✅ No cookies or tracking (unless you add analytics)

## 🐛 Troubleshooting

### Site not loading after deployment
- Wait 2-3 minutes after pushing changes
- Clear browser cache (Ctrl+Shift+R or Cmd+Shift+R)
- Check GitHub Actions tab for build errors

### File processing not working
- Open browser DevTools (F12)
- Check Console tab for errors
- Verify the sheet name is "Existing Comps Data"

### Styles not loading
- Check that `styles.css` is in the same folder as `index.html`
- Verify file was pushed to GitHub (`git status`)

## 📞 Support

If you encounter issues:
1. Check the browser console (F12 → Console)
2. Verify the Excel file has the correct format
3. Make sure the sheet name is "Existing Comps Data"

## 🎓 Learning Resources

- [GitHub Pages Documentation](https://docs.github.com/en/pages)
- [Git Basics](https://git-scm.com/book/en/v2/Getting-Started-Git-Basics)
- [HTML/CSS/JS Basics](https://developer.mozilla.org/en-US/docs/Learn)

---

**You're all set!** Your brother can now use the tool from anywhere with an internet connection. 🎉

