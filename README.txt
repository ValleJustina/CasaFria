Casa Fria website package. Open index.html to view the site.

Included:
- Valle Justina font system: Jost display/interface text and Lora body text
- Latest Casa Fria pool and aerial photos
- Embedded Google Map for Casa Fria Private Resort
- Payment section with the saved QR image at img/gcash-qr.webp
- Existing Apps Script endpoint, exclusive-date availability calendar, booking form, and BOOK NOW CTAs

Calendar rule: because Casa Fria is exclusive-use, any Hold or Booked record returned for a date blocks that entire date. A Booked record takes priority over Hold.

Backend setup:
- Code.gs is the supplied Apps Script updated to use the Casa Fria Google Sheet ID.
- Paste Code.gs into the Sheet's Apps Script project, run setup once, then deploy a new web app version.
- Keep the existing /exec URL in index.html if updating the same deployment.
- Verify the notification email addresses and deposit-slip Drive folder before deployment.
