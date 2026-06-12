# Azure Deployment Steps

## Option 1: Azure Portal Upload / Static Web App Deployment

Use this when IT wants the simplest managed hosting path.

1. Create an Azure Static Web App.
2. Use a deployment source approved by IT, such as GitHub, Azure DevOps, or another internal release process.
3. Set the app location to the folder containing `index.html`.
4. Leave API location blank.
5. Leave build output location blank or set it to the same static app folder, depending on the deployment method.
6. Deploy.
7. Configure Entra ID authentication.
8. Add the company custom domain.

## Option 2: Internal Static Web Server

Use this if the company already has an internal web server.

1. Extract the package.
2. Copy the folder contents to the web root for the app.
3. Ensure `index.html` is the default document.
4. Require company SSO, VPN, or network-based access.
5. Serve over HTTPS.

## Option 3: Azure Blob Static Website

Use this only if access control is handled in front of the site, such as with an internal gateway. Azure Blob static website hosting alone is not the preferred option for this prototype because company authentication matters.

## Browser Support

Use current Microsoft Edge or Chrome. The app uses the browser's local decompression APIs to read `.xlsx` files without sending them to a server.

## Validation Checklist

- App loads over HTTPS.
- Unapproved users cannot access the app.
- Approved users can sign in.
- Workbook can be selected and parsed locally.
- Cockpit metrics render.
- Recommendation queue renders.
- Exports work.
- Network inspection shows no workbook upload.
