// İlke Sendika - Supabase yapılandırması
(function(){
  const SUPABASE_URL = 'https://dulnucqcglytdkiditbw.supabase.co';
  const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImR1bG51Y3FjZ2x5dGRraWRpdGJ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjM0NDczOTUsImV4cCI6MjA3OTAyMzM5NX0.9vRqc2cfE8yXLbSqwskwU33XUIlbRaMJQX7vJy77gR8';

  if (!window.supabase) {
    console.error('Supabase kitaplığı yüklenemedi. Lütfen vendor/supabase/supabase.js dosyasının bu dosyadan önce yüklendiğinden emin olun.');
    return;
  }

  window.ILKE_SUPABASE_URL = SUPABASE_URL;
  window.ilkeSupabase = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
    auth: { persistSession: false }
  });
})();
