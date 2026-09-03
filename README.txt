Casa Fria website package. Open index.html to view the site.

Included:
- Valle Justina font system: Jost display/interface text and Lora body text
- Latest Casa Fria pool and aerial photos
- Embedded Google Map for Casa Fria Private Resort
- Standalone Guest Guide at rules.html, linked from the main navigation and footer
- BDO and GCash details placed directly below the booking total, with the saved QR image at img/gcash-qr.webp
- Existing Apps Script endpoint, multi-select availability calendar, booking form, and BOOK NOW CTAs

Calendar rule: guests can select multiple Package 1, Package 2, and Package 3 cells across multiple dates. Hold and Booked states apply to the matching package and date. Selecting an open cell again removes it.

The website displays peso currency, but the Total written to the Bookings tab is a plain number without a peso symbol.

Backend setup:
- Code.gs is the supplied Apps Script updated to use the Casa Fria Google Sheet ID.
- The setup creates only the Update Here and Bookings tabs.
- Setup and booking notification subjects begin with "CASA FRIA |".
- Paste Code.gs into the Sheet's Apps Script project, run setup once, then deploy a new web app version.
- Keep the existing /exec URL in index.html if updating the same deployment.
- Verify the notification email addresses and deposit-slip Drive folder before deployment.
