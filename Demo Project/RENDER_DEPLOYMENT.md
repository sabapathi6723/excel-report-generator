# Deploying to Render

This Flask application can be deployed as a **Web Service** on Render.

## Prerequisites

1. A GitHub account (sabapathi6723)
2. Your code pushed to a GitHub repository
3. A Render account (sign up at https://render.com)

## Deployment Steps

### Option 1: Using render.yaml (Recommended)

1. **Push your code to GitHub** (if not already done):
   ```bash
   git add .
   git commit -m "Prepare for Render deployment"
   git push origin main
   ```

2. **Go to Render Dashboard**: https://dashboard.render.com

3. **Create New Web Service**:
   - Click "New +" → "Web Service"
   - Connect your GitHub repository
   - Select the repository containing this code

4. **Configure the service**:
   - **Name**: `excel-report-generator` (or your preferred name)
   - **Environment**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
   - **Root Directory**: `Demo Project` (if your app.py is in a subfolder)

5. **Environment Variables** (optional, but recommended):
   - `SECRET_KEY`: Generate a secure random string (Render can auto-generate this)
   - `FLASK_ENV`: `production`

6. **Click "Create Web Service"**

Render will automatically detect `render.yaml` and use those settings.

### Option 2: Manual Configuration

If not using render.yaml, configure manually:

- **Build Command**: `pip install -r requirements.txt`
- **Start Command**: `gunicorn app:app`
- **Environment**: `Python 3`

## Important Notes

1. **File Uploads**: The `uploads` folder is created automatically. On Render, uploaded files are stored temporarily in the filesystem.

2. **Secret Key**: The app uses `SECRET_KEY` environment variable. Render can auto-generate this.

3. **Port**: The app automatically uses Render's `PORT` environment variable (no need to set it manually).

4. **Static Files**: Your templates and static files are included automatically.

## Testing After Deployment

1. Once deployed, Render will give you a URL like: `https://excel-report-generator.onrender.com`

2. Visit the URL and test:
   - Upload a CSV or Excel file
   - Generate a Participation Report
   - Generate a Performance Report

## Troubleshooting

- **Build fails**: Check that all dependencies are in `requirements.txt`
- **App crashes**: Check Render logs (available in dashboard)
- **File upload issues**: Ensure `uploads` folder has write permissions (handled automatically)

## Free Tier Limitations

Render's free tier:
- Services spin down after 15 minutes of inactivity
- First request after spin-down may take 30-60 seconds
- 512MB RAM limit
- 100GB bandwidth/month

For production use, consider upgrading to a paid plan.

