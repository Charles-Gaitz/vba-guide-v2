(function () {
  var header = document.querySelector('.site-header');
  if (!header) return;

  function update() {
    if (window.innerWidth >= 640) {
      header.classList.remove('header-scrolled');
      return;
    }
    if (window.scrollY > 10) {
      header.classList.add('header-scrolled');
    } else {
      header.classList.remove('header-scrolled');
    }
  }

  window.addEventListener('scroll', update, { passive: true });
  window.addEventListener('resize', update);
  update();
})();
