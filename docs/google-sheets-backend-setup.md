# RoverMart Google Sheets backend (starter)

This branch adds a Google Apps Script backend scaffold. It is **not live yet**: you must create the spreadsheet, configure Apps Script properties, deploy the web app, and connect the frontend to its URL. The current website still uses browser localStorage until that integration is completed and tested.

## Architecture

- Google Sheets is the structured order/product database.
- Google Drive stores the spreadsheet privately and can hold product images/invoices.
- Google Apps Script is the API and validates customer input, product prices, stock and variants server-side.
- There is intentionally **no public admin API** in this starter. Manage order statuses in the private spreadsheet. Editing the status column on the Orders sheet records a tracking-history entry through an installable trigger.

## 1. Create the spreadsheet

1. Create a new Google Spreadsheet in the Google account that will own RoverMart.
2. Keep its sharing set to **Restricted**. Do not publish it or grant public access.
3. Copy the spreadsheet ID from its URL: `https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit`.
4. Open **Extensions → Apps Script**.

## 2. Add the backend code

1. Replace the starter `Code.gs` content with the code in [Code.gs](./Code.gs), or paste it into the Apps Script editor.
2. In Apps Script **Project Settings → Script Properties**, add:
   - Property: `SPREADSHEET_ID`
   - Value: the spreadsheet ID from step 3.
3. In **Project Settings**, set the time zone to `Asia/Dhaka`.
4. Save. Select `setupSheets` and click **Run** once. Review and approve the Google permissions prompt. This creates the tabs and installs the status-edit trigger.
5. In the `Products` tab, add product rows. Use `TRUE` in `active`; enter variants separated by `|` (example: `S|M|L`). Product IDs must be unique and stable.

## 3. Deploy as a web app

1. Select **Deploy → New deployment → Web app**.
2. Execute as: **Me** (the spreadsheet owner).
3. Access: **Anyone** so customers can place and track orders without being granted access to the private spreadsheet.
4. Deploy and copy the Web App URL ending in `/exec`.
5. Do not paste credentials, service tokens, or spreadsheet IDs into public frontend code. The spreadsheet ID remains only in Script Properties.

Because this public endpoint executes as the owner, the code intentionally exposes only active products, order creation, and tracking protected by a long random token. It does not expose customer lists or admin mutations. Review Google Apps Script quotas and add abuse monitoring before advertising the store widely.

## 4. API contract

### List active products

`GET /exec?action=products`

Returns active products only. Product sheet columns:

`product_id | name | category | price | stock | active | image_url | variants`

### Create order

Send a JSON POST body (use `Content-Type: text/plain;charset=utf-8` to avoid a browser preflight where possible):

```json
{
  "action": "createOrder",
  "customer": {
    "name": "Customer Name",
    "phone": "01XXXXXXXXX",
    "address": "Full delivery address"
  },
  "payment": {
    "method": "Cash on Delivery",
    "senderNumber": "",
    "transactionId": ""
  },
  "items": [
    { "id": "SCARF-001", "quantity": 2, "variant": "Standard" }
  ]
}
```

For bKash/Nagad, use `method` as `bKash` or `Nagad` and provide the sender number and transaction reference. This is only manual payment-reference collection, not payment-gateway verification.

Successful response includes `orderId` and a `trackingToken`. **Show/save the token for the customer once**. The backend stores only a SHA-256 hash of it. The order ID alone cannot reveal order details.

### Track an order

`GET /exec?action=track&orderId=RM-YYYYMMDD-XXXXXXXX&token=TRACKING_TOKEN`

Only returns status, items, totals and tracking history. It does not return the customer's name, phone, address, payment sender number or transaction reference. Treat the tracking token like a password; do not put it in public posts or analytics.

## 5. Admin workflow for the first release

1. Open the restricted spreadsheet as an authorized shop admin.
2. Review new orders in `Orders`, `OrderItems` and `Payments`.
3. Change `Orders.status` (column G) to one of:
   - `Pending`
   - `Payment Verified`
   - `Payment Not Verified`
   - `Processing`
   - `Completed`
   - `Cancelled`
4. The installable trigger appends the status update to `OrderTracking`. Avoid bulk-pasting statuses unless you also verify the resulting history entries.
5. Update payment verification separately in `Payments.verification_status`.

## Important limitations before launch

- The existing `index.html` is not yet wired to this API. Browser localStorage remains the current behavior until the frontend integration is implemented.
- Existing localStorage orders are per browser and cannot be migrated automatically by this server. Export them manually only if needed; review personal data before importing.
- The backend scaffold is not a substitute for a full security review. Add rate limiting/abuse monitoring and test concurrency, quotas, malformed requests, and error handling before public launch.
- Admin authentication is not implemented as a website login. The spreadsheet is the initial private admin console; do not re-enable the website admin panel until a proper server-verified admin authentication flow exists.
- Google Sheets is suitable for a small, low-volume shop but is not a transactional database at scale.
