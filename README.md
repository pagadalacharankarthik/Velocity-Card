# Velocity Card Engine

A static web application prototype for creating and managing premium digital wedding invitations. It includes a storefront, order form, invitation templates, a shop portal, and an admin dashboard. The project explores how print shops can manage digital wedding invitation orders and invitation designs through a simple web interface.

## Features

* **Storefront:** Landing page showcasing invitation designs.
* **Order Form:** Interface for collecting wedding invitation details.
* **Invitation Templates:** Designs for digital wedding invitations.
* **Shop Portal:** Interface for shop-related activities.
* **Admin Dashboard:** Pages for managing orders and shops.
* **Prototype Backend:** Optional Google Apps Script integration using Google Sheets.

## Project Structure

```text
Velocity-Card-Engine/
├── index.html
├── order.html
├── shop.html
├── admin/
│   ├── login.html
│   ├── order.html
│   └── shop.html
├── templates/
├── code.gs
└── README.md
```

* `index.html` — Landing page and invitation showcase.
* `order.html` — Invitation order form.
* `shop.html` — Shop portal interface.
* `admin/` — Admin login and management pages.
* `templates/` — Invitation design templates.
* `code.gs` — Optional Google Apps Script backend using Google Sheets.

## Run Locally

The project uses HTML, CSS, and JavaScript and can be previewed with a local HTTP server.

1. Install [Node.js](https://nodejs.org) if it is not already installed.
2. Open a terminal in the project directory.
3. Run:
   ```bash
   npx --yes http-server . -p 4173
   ```
4. Open [http://127.0.0.1:4173](http://127.0.0.1:4173) in your browser.

Press `Ctrl+C` in the terminal to stop the server.

## Technology Stack

* HTML, CSS, and JavaScript
* Google Apps Script (optional backend)
* Google Sheets (optional data storage)
* Browser libraries loaded through CDNs, where applicable

## Project Status

**Status: Prototype — Not Yet Deployed**

This project is currently a basic prototype intended to explore the user interface and workflow for digital wedding invitation management.

* The front-end pages can be previewed locally.
* Backend integration requires additional configuration and testing.
* Order processing, authentication, data persistence, and admin permissions may not be fully functional.
* Production deployment and security hardening have not been completed.

## Future Improvements

* Complete and test backend integration.
* Add secure authentication and role-based access control.
* Store orders and invitation details in a reliable database.
* Improve invitation customization and preview features.
* Validate forms and add error handling.
* Deploy the application after testing the complete workflow.

## Disclaimer

This project is a work in progress. Features shown in the interface may require additional implementation before they can be used in a production environment.
