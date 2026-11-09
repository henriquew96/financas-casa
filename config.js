// Conexão com a nuvem (Supabase) — preencher com os dados do projeto pessoal.
// Supabase → Project Settings → API (ou Data API): "Project URL" e a chave "publishable" (sb_publishable_...) ou "anon".
// Essas duas informações são públicas por natureza: quem protege os dados é o login + as regras do banco (RLS, em banco.sql).
// Deixando vazio, o app funciona só no aparelho (sem login e sem sincronizar).
window.FINANCAS_CONFIG = {
  supabaseUrl: 'https://ypzxjopuuwhsleztbatu.supabase.co',
  supabaseChave: 'sb_publishable_6lFYQYUcSb-UHQeG10UI-A_akd7J--l'
};
