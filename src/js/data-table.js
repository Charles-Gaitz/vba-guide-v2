document.querySelectorAll('.data-table-toggle').forEach(function (btn) {
  btn.addEventListener('click', function () {
    var wrap = this.parentElement.querySelector('.data-table-wrap');
    if (!wrap) return;
    var nowVisible = wrap.hidden;
    wrap.hidden = !nowVisible;
    this.setAttribute('aria-expanded', String(nowVisible));
    this.textContent = nowVisible ? 'Hide Data Table' : '📋 Show Data Table';
  });
});

document.querySelectorAll('.copy-data-btn').forEach(function (btn) {
  btn.addEventListener('click', function () {
    var tsv = this.getAttribute('data-tsv');
    if (!tsv) return;
    var self = this;
    navigator.clipboard.writeText(tsv).then(function () {
      self.textContent = '✓ Copied!';
      setTimeout(function () { self.textContent = 'Copy to Clipboard'; }, 2000);
    }).catch(function () {
      self.textContent = 'Copy failed — select the table manually';
      setTimeout(function () { self.textContent = 'Copy to Clipboard'; }, 2000);
    });
  });
});
