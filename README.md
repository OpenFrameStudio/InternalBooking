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

## Files

- `server.js` serves the app, validates bookings, stores data, and talks to Lark.
- `public/` contains the booking interface.
- `data/bookings.json` stores local bookings.
- `data/clients.json` stores reusable client and agent details.
