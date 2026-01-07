# Content Security Policy (CSP)

This project ships with a production‑ready CSP meta tag in pages and a small script that relaxes the policy on localhost (development) to simplify testing.

## Production CSP

The default CSP allows Google reCAPTCHA and Supabase:

```
<meta http-equiv="Content-Security-Policy" id="csp-meta" content="default-src 'self'; script-src 'self' https://www.google.com https://www.gstatic.com 'unsafe-inline'; style-src 'self' 'unsafe-inline' https://www.gstatic.com; img-src 'self' data: https://www.google.com https://www.gstatic.com https://*.supabase.co; font-src 'self' data:; connect-src 'self' https://www.google.com https://www.gstatic.com https://*.supabase.co; frame-src https://www.google.com https://www.gstatic.com; base-uri 'self'; form-action 'self'" />
```

If you deploy behind a server (NGINX/Netlify/Cloudflare), prefer sending CSP via response headers instead of meta tags. Use the same directives as above.

## Localhost relax (development)

On `localhost` or `127.0.0.1`, a small script detects the hostname and replaces the CSP with a more permissive one:

```
<script>
(function(){
  try{
    var isLocal = /^(localhost|127\.0\.0\.1)$/i.test(location.hostname);
    if (!isLocal) return;
    var m = document.getElementById('csp-meta');
    if (!m) return;
    var relaxed = "default-src 'self'; script-src 'self' https://www.google.com https://www.gstatic.com 'unsafe-inline' 'unsafe-eval'; style-src 'self' 'unsafe-inline' https://www.gstatic.com; img-src 'self' data: https://www.google.com https://www.gstatic.com https://*.supabase.co; font-src 'self' data:; connect-src *; frame-src https://www.google.com https://www.gstatic.com; base-uri 'self'; form-action 'self'";
    m.setAttribute('content', relaxed);
  }catch(e){}
})();
</script>
```

This avoids noisy console warnings on localhost while keeping production strict.

## Server header configuration (optional)

If you manage headers at the server/CDN level, add a `Content-Security-Policy` header equivalent to the production policy above. Example NGINX snippet:

```
add_header Content-Security-Policy "default-src 'self'; script-src 'self' https://www.google.com https://www.gstatic.com 'unsafe-inline'; style-src 'self' 'unsafe-inline' https://www.gstatic.com; img-src 'self' data: https://www.google.com https://www.gstatic.com https://*.supabase.co; font-src 'self' data:; connect-src 'self' https://www.google.com https://www.gstatic.com https://*.supabase.co; frame-src https://www.google.com https://www.gstatic.com; base-uri 'self'; form-action 'self'" always;
```

Adjust to your own origins (custom CDNs, analytics, etc.).
# İlke Sendika Web Sitesi

Modern, hızlı ve erişilebilir statik web sitesi taslağı. Harici CDN kullanılmaz; tüm varlıklar yerelden yüklenir.

## Yapı

- `index.html` — Ana sayfa (Hero, Hakkımızda, Haberler, Duyurular, İletişim)
- `styles.css` — Minimal, modern ve responsive stil dosyası
- `app.js` — Basit etkileşimler (menü, yıl, iletişim formu doğrulama)
- `assets/logo.svg` — Logo

## Geliştirme

1) Dosyayı tarayıcıda açın:
   - `index.html` dosyasını çift tıklayarak veya VS Code Live Server ile açabilirsiniz.

2) Yerel ön izleme (opsiyonel):
   - Python 3 ile:
     ```bash
     python -m http.server 5500
     ```
     Ardından: http://localhost:5500/ilke-sendika/

   - Node.js ile (npx serve):
     ```bash
     npx serve -l 5500
     ```

## Dağıtım (Deployment)

- GitHub Pages / Netlify / Vercel gibi statik hosting servisleri ile doğrudan yayınlanabilir.
- Kök dizin: `ilke-sendika/`

## Özelleştirme

- Renkler: `styles.css` üzerinde `:root` değişkenlerini güncelleyebilirsiniz.
- Metinler: `index.html` içindeki bölümleri (Hakkımızda, Haberler, Duyurular, İletişim) güncelleyin.
- Logo: `assets/logo.svg` dosyasını değiştirin.

## Erişilebilirlik ve SEO

- `skip-link`, semantik başlıklar, `aria-` öznitelikleri kullanıldı.
- Temel Open Graph meta etiketleri ve açıklama başlığında mevcuttur.

## Lisans

Tüm hakları İlke Sendika'ya aittir. (Tasnif: İç kullanım)
