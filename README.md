# Lark Booking Desk

A small booking system that stores appointments locally and can sync confirmed bookings to a Lark Calendar.

Current admin-selectable services:

- Photography
- Floorplan
- Drone
- Siteplan

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
- Test employee login: `Test`
- Test employee password: `1111`

Faye and Test can access Bookings and Work. Boss can access Bookings, Clients, Photographers, Invoices, Wages, and Work.
Work actions are enforced by server-side permissions; hiding buttons in the browser is only for usability.

Work assignments include an assignee dropdown. Faye is the default employee and the default assignee for imported booking work; Test is available as a quiet testing employee account.

You can override these with `ADMIN_USERNAME`, `ADMIN_PASSWORD`, `EMPLOYEE_USERNAME`, `EMPLOYEE_PASSWORD`, `TEST_EMPLOYEE_USERNAME`, and `TEST_EMPLOYEE_PASSWORD`.

Invoices are created automatically from confirmed bookings. The default draft invoice prices are:

- Photography: `$150`
- Floorplan: `$75`
- Drone: `$100`
- Siteplan: `$25`

Invoices always add 10% GST to the service subtotal.

Wages are split into two sheets. Employee wages are for full-time or part-time staff, with Faye seeded as a monthly `THB 15,000` employee wage paid on the last day of every month. Contractor wages are photographer proformas created automatically from confirmed bookings; photography-only is `$90`, and each photographer can be marked as GST included or no GST.

## Deploy to openframe.studio

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
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_ROLE_KEY`
   - optional: `LARK_ORGANIZER_CALENDAR_ID` for the admin sender calendar
   - optional: `LARK_SENDER_EMAIL` (defaults to `admin@openframe.studio`)
   - optional: `LARK_SENDER_NAME` (defaults to `admin@openframe.studio`)
   - `RESEND_API_KEY`
   - optional: `APP_PUBLIC_URL` (defaults to `https://system.openframe.studio`)
   - optional: `INVOICE_EMAIL_PROVIDER` (defaults to `resend`)
   - optional: `INVOICE_EMAIL_FROM` (defaults to `OpenFrame Studio <admin@openframe.studio>`)
   - optional: `WORK_INVITE_EMAIL_TO` for Faye's work assignment invite emails
   - optional: `WORK_INVITE_EMAIL_FROM` (defaults to `OpenFrame Studio <admin@openframe.studio>`)
   - optional: `WORK_LARK_RECEIVE_ID` for Faye's Lark work notifications
   - optional: `WORK_LARK_RECEIVE_ID_TYPE` (defaults to `email`)
   - optional: `WORK_COMPLETION_LARK_RECEIVE_ID` for Barry's completion notifications (defaults to `barry.gao@openframe.studio`)
   - optional: `WORK_COMPLETION_LARK_RECEIVE_ID_TYPE` (defaults to `email`)
   - optional: `ADMIN_USERNAME`
   - optional: `ADMIN_PASSWORD`
4. After the service is live, open its Custom Domains settings and verify
   `internalbooking.openframe.studio` and `system.openframe.studio`.
5. In your DNS provider, add the DNS records Render gives you for the
   `internalbooking` and `system` subdomains, then verify them in Render.

The blueprint sets the app to listen on `0.0.0.0` in production and uses
`/api/status` as the health check.

## Supabase storage

Supabase is the preferred permanent storage for bookings, clients, photographers, invoices, wages, work items, and changed passwords.

1. Create a Supabase project.
2. Open SQL Editor and run `supabase-schema.sql`.
3. In Render, add:
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_ROLE_KEY`
   - `SUPABASE_TABLE=internal_booking_data`
4. Deploy the app and check `/api/status`; `storageBackend` should be `supabase`.

The service role key must stay server-side in Render only. Do not use the anon key for this app.

GitHub storage was removed after the Supabase migration. Remove any old
`GITHUB_STORAGE_*` variables from Render; they are no longer used by the app.

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

Bookings are created without a Lark video meeting by sending `vchat.vc_type`
as `no_meeting`.

## Send Invoices

The invoice desk can send a PDF tax invoice to the saved client email through
Resend over HTTPS, which works on Render free.

```text
INVOICE_EMAIL_PROVIDER=resend
RESEND_API_KEY
INVOICE_EMAIL_FROM=OpenFrame Studio <admin@openframe.studio>
INVOICE_EMAIL_REPLY_TO=admin@openframe.studio
```

Verify `openframe.studio` in Resend first so emails can send from
`admin@openframe.studio`.

Work assignment invites also use Resend over HTTPS. Set
`WORK_INVITE_EMAIL_TO` to Faye's email address; the default sender is
`OpenFrame Studio <admin@openframe.studio>`.

Work assignment Lark notifications use your existing Lark app to send a bot
message to Faye when new work is assigned. In Lark Open Platform, enable the
bot and grant permission to send messages as the app. The app can send to an
email recipient, then fall back to looking up Faye's `open_id` by email; that
fallback needs Lark's contact user ID lookup permission. If Lark still rejects
the email, set `WORK_LARK_RECEIVE_ID` to Faye's Lark user identifier and change
`WORK_LARK_RECEIVE_ID_TYPE` to `open_id`, `user_id`, `union_id`, or `chat_id`.

When an employee marks work finished, the app sends Barry a Lark notification
at `barry.gao@openframe.studio` by default. Override it with
`WORK_COMPLETION_LARK_RECEIVE_ID` if needed.

SMTP is still available if `INVOICE_EMAIL_PROVIDER=smtp`, but Render free
services block outbound SMTP ports such as `465` and `587`.

## Files

- `server.js` serves the app, validates bookings, stores data, and talks to Lark.
- `public/` contains the booking interface.
- `data/bookings.json` stores local bookings.
- `data/clients.json` stores reusable client and agent details.
