// WarungQ - script umum
document.addEventListener('DOMContentLoaded', function () {
  // Tombol mata: intip password
  document.querySelectorAll('.pw-wrap').forEach(function (wrap) {
    var input = wrap.querySelector('input');
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'eye-btn';
    btn.innerHTML = '👁';
    btn.title = 'Intip password';
    btn.setAttribute('aria-label', 'Tampilkan password');
    btn.addEventListener('click', function () {
      if (input.type === 'password') { input.type = 'text'; btn.innerHTML = '🙈'; }
      else { input.type = 'password'; btn.innerHTML = '👁'; }
    });
    wrap.appendChild(btn);
  });

  // Konfirmasi hapus
  document.querySelectorAll('form[data-confirm]').forEach(function (f) {
    f.addEventListener('submit', function (e) {
      if (!confirm(f.getAttribute('data-confirm'))) e.preventDefault();
    });
  });
});

// Format rupiah
function rp(n) { return 'Rp' + Number(n || 0).toLocaleString('id-ID'); }
