# Velocity Card Engine

A static web application for creating and managing premium digital wedding invitations. It includes a storefront, order form, invitation templates, a shop portal, and an admin dashboard.

## Pages

- `index.html` - Product landing page and template showcase.
- `order.html` - Invitation order form.
- `shop.html` - Shop account portal.
- `admin/login.html` - Admin dashboard and login.
- `admin/order.html` and `admin/shop.html` - Admin order and shop views.
- `templates/` - Invitation designs used for published invitations.
- `supabase-app.js` - Local Supabase integration; intentionally excluded from this public repository.
- `code.gs` - Optional Google Apps Script backend using Google Sheets.

## Run locally

The front end is static and can be served using any local HTTP server. With Node.js installed, run this from the project directory:

```powershell
npx --yes http-server . -p 4173 -c-1
```

Then open <http://127.0.0.1:4173>. Stop the server with `Ctrl+C`.

Opening the HTML files directly may work for simple previews, but a local HTTP server is recommended for consistent browser behavior and relative links.

## Configuration

The public repository does not include `supabase-app.js`, so Supabase-backed order, account, upload, RSVP, analytics, and admin workflows are not functional from this checkout until an integration is supplied. The HTML pages still contain script references to that excluded file. The optional Apps Script backend also needs configuration and deployment.

1. In `code.gs`, replace `YOUR_GOOGLE_SHEET_ID` with the spreadsheet ID and set the `ADMIN_PASSWORD` script property in Apps Script project settings.
2. Deploy the Apps Script project as a web app and connect the front end to the deployed endpoint as required by the app.
3. Configure optional Cloudinary values if image hosting is used.

Never put passwords, private keys, or service-role credentials in browser code or a public repository. Rotate any credentials that have already been exposed. This project has no package manifest or automated build step; its pages load browser libraries such as GSAP and the Supabase client from CDNs.

## Deployment

Deploy the static files to a static hosting provider and configure the required backend services separately. Keep credentials and environment-specific configuration out of the public client bundle. Invitation and admin flows may require additional backend setup before deployment.
