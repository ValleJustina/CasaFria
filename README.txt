CASA FRIA WEBSITE

1. Upload index.html and the assets folder to your web host.
2. Create a Google Sheet and open Extensions > Apps Script.
3. Paste Code.gs into the Apps Script project.
4. Deploy as a Web App with access configured for your intended public booking flow.
5. Copy the Web App /exec URL into CONFIG.endpoint in index.html.
6. The first reservation submission creates a Reservations sheet automatically.

Important: the backend currently allows one active reservation request per date. If Casa Fria can host more than one booking on the same date, change the conflict rule before deployment.

Contact numbers, email, payment instructions, logo, social links, and exact address were not supplied, so they were intentionally not invented.
