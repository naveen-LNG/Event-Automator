# Event Automation Tool - GitHub Deployment Script

Write-Host "--- Event Automation Tool: GitHub Deployment ---" -ForegroundColor Cyan

# 1. Initialize Git if not already present
if (-not (Test-Path ".git")) {
    Write-Host "Initializing new Git repository..."
    git init
}

# 2. Add all files
git add .

# 3. Initial commit
$commitMsg = "Initial commit: Finalized 21-step event automation tool"
git commit -m "$commitMsg"

# 4. Prompt for Remote URL
Write-Host ""
$remoteUrl = Read-Host "Please enter your empty GitHub Repository URL (e.g., https://github.com/user/repo.git)"

if (-not [string]::IsNullOrWhiteSpace($remoteUrl)) {
    # Check if remote already exists, if so remove then add
    $remotes = git remote
    if ($remotes -contains "origin") {
        git remote remove origin
    }
    
    git remote add origin $remoteUrl
    
    Write-Host "Pushing to GitHub (Main branch)..."
    git branch -M main
    git push -u origin main
    
    Write-Host ""
    Write-Host "--- DEPLOYMENT SUCCESSFUL! ---" -ForegroundColor Green
    Write-Host "Your tool is now available on GitHub."
} else {
    Write-Host "No URL provided. Deployment cancelled. Your files are committed locally." -ForegroundColor Yellow
}

Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
