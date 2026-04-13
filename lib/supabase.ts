import { createClient } from "@supabase/supabase-js";

export const supabase = createClient(
  "https://YOUR_PROJECT_ID.supabase.co",
  "YOUR_ANON_PUBLIC_KEY"
);