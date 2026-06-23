// Copie este arquivo para config.js e preencha com suas credenciais.
// config.js está no .gitignore — nunca commite seus valores reais.
//
// Supabase: https://supabase.com/dashboard → seu projeto → Settings → API
window.APP_CONFIG = {
  SUPABASE_URL: "https://COLE_AQUI.supabase.co",
  SUPABASE_ANON_KEY: "COLE_AQUI",
  COMPANY_NAME: "Minha Empresa",
  ADMIN_PASSWORD: "COLE_AQUI",
};

window.IS_SUPABASE_CONFIGURED =
  window.APP_CONFIG.SUPABASE_URL &&
  !window.APP_CONFIG.SUPABASE_URL.includes("COLE_AQUI") &&
  window.APP_CONFIG.SUPABASE_ANON_KEY &&
  !window.APP_CONFIG.SUPABASE_ANON_KEY.includes("COLE_AQUI");
