(function(){
  try{
    // Safe Images configuration for Admin Panel
    window.__SAFE_IMG_WRAP_ALL = true;
    window.__SAFE_IMG_HOSTS = ['supabase.co'];
    // Do not wrap only google charts; allow QR server to be wrapped as blob to avoid canvas taint
    window.__SAFE_IMG_SKIP_HOSTS = ['chart.googleapis.com'];
  }catch(_){}
})();
