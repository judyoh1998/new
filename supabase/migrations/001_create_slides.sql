create table slides (
  id uuid primary key default gen_random_uuid(),
  title text not null,
  description text not null,
  content text not null,
  created_at timestamptz default now()
);
