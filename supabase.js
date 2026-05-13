import { createClient } from "@supabase/supabase-js";

const supabaseUrl = "https://yhutgxfvxvmvyttkcwja.supabase.co";

const supabaseKey =
  "sb_publishable_mekbpTaPESea_kj7W42TKA_ckHWyvQG";

export const supabase = createClient(supabaseUrl, supabaseKey);
