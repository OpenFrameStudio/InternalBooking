# Lark Booking Desk

A small booking system that stores appointments locally and can sync confirmed bookings to a Lark Calendar.

Current admin-selectable services:

- Photography
- Floorplan
- Drone

Calendar entries use one property address for the event title, Lark location name, and Lark full address. Notes are formatted with the agency/client, photographer, agent, selected services, and any extra instructions. Guest emails are saved locally and, when Lark is configured, sent to Lark as third-party attendees.

Saved clients live in `data/clients.json` and are managed on the Clients page. Selecting a saved client on the Bookings page fills the client email, agent name, and agent phone. The client email is automatically included in the booking guest email list.

## Run it

```bash
node server.js
```

Open `http://127.0.0.1:4180`.

Default logins:

- Boss login: `ShuhanGao`
- Boss password: `Sg1654723576`
- Faye login: `Faye`
- Faye password: `1111`

Faye can access Bookings and Work. Boss can access Bookings, Clients, Photographers, and Work.

You can override these with `ADMIN_USERNAME`, `ADMIN_PASSWORD`, `EMPLOYEE_USERNAME`, and `EMPLOYEE_PASSWORD`.

Invoices are created automatically from confirmed bookings. The default draft invoice prices are:

- Photography: `$150`
- Floorplan: `$50`
- Drone: `$50`

Invoices always add 10% GST to the service subtotal.

## Deploy to internalbooking.openframe.studio

This repo includes `render.yaml` for deploying the app as a Render web service.
The default blueprint uses Render's free web service so you can get the app live
first. On the free plan, saved bookings and clients can be reset when Render
restarts or redeploys the service.

1. In Render, create a new Blueprint from this GitHub repo.
2. Keep the service as `internal-booking`.
3. Fill in the prompted secret environment variables:
   - `LARK_APP_ID`
   - `LARK_APP_SECRET`
   - `LARK_CALENDAR_ID`
   - optional: `LARK_ORGANIZER_CALENDAR_ID` for the admin sender calendar
   - optional: `LARK_SENDER_EMAIL` (defaults to `admin@openframe.studio`)
   - optional: `LARK_SENDER_NAME` (defaults to `admin@openframe.studio`)
   - `INVOICE_EMAIL_HOST`
   - `INVOICE_EMAIL_USER`
   - `INVOICE_EMAIL_PASSWORD`
   - optional: `INVOICE_EMAIL_PORT` (defaults to `587`)
   - optional: `INVOICE_EMAIL_FROM` (defaults to the email user)
   - optional: `ADMIN_USERNAME`
   - optional: `ADMIN_PASSWORD`
4. After the service is live, open its Custom Domains settings and verify
   `internalbooking.openframe.studio`.
5. In your DNS provider, add the DNS record Render gives you for the
   `internalbooking` subdomain, then verify it in Render.

The blueprint sets the app to listen on `0.0.0.0` in production and uses
`/api/status` as the health check. Upgrade the service before using it for real
bookings so the saved bookings and clients can be moved to permanent storage.

## Connect Lark Calendar

1. Create a self-built app in the Lark Open Platform.
2. Grant the app Calendar permissions for creating events.
3. Add the app to your Lark tenant.
4. Copy `.env.example` to `.env`, fill in your Lark values, then start the app:

```bash
node server.js
```

The app calls Lark's tenant token endpoint and then creates events at:

```text
POST /open-apis/auth/v3/tenant_access_token/internal
POST /open-apis/calendar/v4/calendars/{calendar_id}/events
```

For invite emails to show from the admin account, set
`LARK_ORGANIZER_CALENDAR_ID` to the actual Lark calendar ID owned by
`admin@openframe.studio`. Lark controls the final email sender header, so the
calendar that owns the event is what matters.

## Send Invoices

The invoice desk can send a PDF tax invoice to the saved client email. Add SMTP
settings in Render for an email account that is allowed to send from
`admin@openframe.studio`:

```text
INVOICE_EMAIL_HOST
INVOICE_EMAIL_PORT
INVOICE_EMAIL_SECURE
INVOICE_EMAIL_USER
INVOICE_EMAIL_PASSWORD
INVOICE_EMAIL_FROM
INVOICE_EMAIL_REPLY_TO
```

## Files

- `server.js` serves the app, validates bookings, stores data, and talks to Lark.
- `public/` contains the booking interface.
- `data/bookings.json` stores local bookings.
- `data/clients.json` stores reusable client and agent details.
