
import { createClient } from '@supabase/supabase-js';

const cleanEnvVar = (val: string | undefined): string => {
  if (!val) return '';
  return val.trim().replace(/^["']|["']$/g, '');
};

const supabaseUrl = cleanEnvVar((import.meta as any).env.VITE_SUPABASE_URL);
const supabaseAnonKey = cleanEnvVar((import.meta as any).env.VITE_SUPABASE_ANON_KEY);

if (!supabaseUrl || !supabaseAnonKey) {
  console.error("KRITIKAL: Supabase URL atau Anon Key TIDAK ditemukan di environment variables (VITE_...). Request API akan gagal (401).");
}

export const supabase = createClient(
  supabaseUrl || 'https://placeholder.supabase.co', 
  supabaseAnonKey || 'placeholder'
);
