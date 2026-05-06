document.querySelectorAll('.ai-compare').forEach(function (compare) {
  const tabs = compare.querySelector('.ai-compare-tabs');
  // No tabs in DOM — panels are always visible, nothing to do
  if (!tabs) return;

  const panels = compare.querySelectorAll('.ai-panel');

  function apply() {
    // Tabs hidden via CSS (display:none) — keep all panels visible and exit
    if (getComputedStyle(tabs).display === 'none') {
      panels.forEach(function (p) { p.hidden = false; });
      return;
    }

    const isMobile = window.innerWidth < 640;
    tabs.hidden = !isMobile;
    if (!isMobile) {
      panels.forEach(function (p) { p.hidden = false; });
    } else {
      const active = compare.querySelector('.ai-tab.active') || tabs.querySelector('.ai-tab');
      const activePanel = active ? active.dataset.panel : 'sanders';
      panels.forEach(function (p) { p.hidden = p.dataset.panel !== activePanel; });
    }
  }

  tabs.addEventListener('click', function (e) {
    const tab = e.target.closest('.ai-tab');
    if (!tab) return;
    compare.querySelectorAll('.ai-tab').forEach(function (t) { t.classList.remove('active'); });
    tab.classList.add('active');
    apply();
  });

  apply();
  window.addEventListener('resize', apply);
});
