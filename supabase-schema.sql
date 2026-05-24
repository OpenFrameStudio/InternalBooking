create table if not exists public.internal_booking_data (
  namespace text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);

alter table public.internal_booking_data enable row level security;

revoke all on table public.internal_booking_data from anon;
revoke all on table public.internal_booking_data from authenticated;
